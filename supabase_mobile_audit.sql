-- Ejecutar UNA vez en Supabase > SQL Editor.
-- Todas estas tablas son privadas: no se concede acceso directo al navegador.
create table if not exists public.mobile_audit_orders (
    id text primary key check (length(id) = 64),
    batch_id text not null,
    pedido text not null,
    slot text not null default '',
    owner_emails text[] not null check (cardinality(owner_emails) > 0),
    picker_labels text not null default '',
    items jsonb not null check (jsonb_typeof(items) = 'array'),
    published_by_email text not null,
    created_at timestamptz not null default now()
);

create table if not exists public.mobile_audit_assignments (
    id uuid primary key default gen_random_uuid(),
    order_id text not null references public.mobile_audit_orders(id),
    mode text not null check (mode in ('autoverificacion','auditoria_cruzada')),
    assignee_email text not null,
    assigned_by_email text not null,
    status text not null default 'pendiente' check (status in ('pendiente','validado')),
    last_result text not null default '',
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique(order_id, mode, assignee_email)
);

create table if not exists public.mobile_audit_reviews (
    id uuid primary key default gen_random_uuid(),
    assignment_id uuid not null references public.mobile_audit_assignments(id),
    order_id text not null references public.mobile_audit_orders(id),
    reviewer_email text not null,
    mode text not null check (mode in ('autoverificacion','auditoria_cruzada')),
    result text not null check (result in ('Pedido correcto','Con diferencias','No se pudo validar')),
    notes text not null default '',
    differences jsonb not null default '[]'::jsonb check (jsonb_typeof(differences) = 'array'),
    checked_items integer not null default 0 check (checked_items >= 0),
    idempotency_key text not null,
    created_at timestamptz not null default now(),
    unique(assignment_id, idempotency_key)
);

create index if not exists mobile_audit_assignments_email_idx
    on public.mobile_audit_assignments (assignee_email, created_at desc);
create index if not exists mobile_audit_reviews_order_idx
    on public.mobile_audit_reviews (order_id, created_at desc);

-- Validación en BD: un picker jamás puede recibir como auditoría cruzada
-- un pedido del que es propietario, aunque falle la interfaz de asignación.
create or replace function public.mobile_audit_check_assignment()
returns trigger
language plpgsql
set search_path = public, pg_temp
as $$
declare
    v_owners text[];
begin
    select owner_emails into v_owners
      from public.mobile_audit_orders
     where id = new.order_id;
    if v_owners is null then
        raise exception 'Pedido no publicado.';
    end if;
    if new.assignee_email is null or
       new.assignee_email <> lower(btrim(new.assignee_email)) or
       position('@' in new.assignee_email) < 2 then
        raise exception 'Correo del responsable inválido.';
    end if;
    if new.mode = 'autoverificacion' and not (new.assignee_email = any(v_owners)) then
        raise exception 'Sólo el picker propietario puede autoverificar.';
    end if;
    if new.mode = 'auditoria_cruzada' and new.assignee_email = any(v_owners) then
        raise exception 'El propietario no puede hacer auditoría cruzada de su pedido.';
    end if;
    return new;
end;
$$;

drop trigger if exists mobile_audit_assignment_guard on public.mobile_audit_assignments;
create trigger mobile_audit_assignment_guard
before insert or update of mode, assignee_email, order_id
on public.mobile_audit_assignments
for each row execute function public.mobile_audit_check_assignment();

-- Registrar resultado y cerrar asignación en una sola transacción,
-- verificando nuevamente quién está autorizado y evitando duplicados.
create or replace function public.submit_mobile_audit_review(
    p_assignment_id uuid,
    p_reviewer_email text,
    p_result text,
    p_notes text,
    p_differences jsonb,
    p_checked_items integer,
    p_idempotency_key text
) returns text
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
    v_assignment public.mobile_audit_assignments%rowtype;
    v_existing uuid;
    v_id uuid;
    v_notes text := btrim(coalesce(p_notes, ''));
    v_count integer := 0;
begin
    select * into v_assignment
      from public.mobile_audit_assignments
     where id = p_assignment_id
     for update;

    if not found then
        raise exception 'Asignación inexistente.';
    end if;
    if lower(btrim(coalesce(p_reviewer_email, ''))) <> v_assignment.assignee_email then
        raise exception 'Esta asignación pertenece a otro usuario.';
    end if;
    if nullif(btrim(coalesce(p_idempotency_key, '')), '') is null then
        raise exception 'Falta identificador único de la revisión.';
    end if;

    select id into v_existing
      from public.mobile_audit_reviews
     where assignment_id = p_assignment_id
       and idempotency_key = p_idempotency_key;
    if v_existing is not null then
        return v_existing::text;
    end if;
    if v_assignment.status <> 'pendiente' then
        raise exception 'Este pedido ya fue validado.';
    end if;

    if p_result not in ('Pedido correcto', 'Con diferencias', 'No se pudo validar') then
        raise exception 'Resultado no permitido.';
    end if;
    if p_checked_items is null or p_checked_items < 0 then
        raise exception 'Cantidad de artículos inválida.';
    end if;
    if p_differences is null or jsonb_typeof(p_differences) <> 'array' then
        raise exception 'Las diferencias deben ser una lista.';
    end if;
    v_count := jsonb_array_length(p_differences);

    if p_result = 'Pedido correcto' and v_count > 0 then
        raise exception 'Un pedido correcto no puede tener diferencias.';
    end if;
    if p_result = 'Con diferencias' and (v_count < 1 or v_notes = '') then
        raise exception 'Describe al menos un artículo y la diferencia.';
    end if;
    if p_result = 'No se pudo validar' and v_notes = '' then
        raise exception 'Explica por qué quedó pendiente.';
    end if;

    insert into public.mobile_audit_reviews(
        assignment_id,order_id,reviewer_email,mode,result,notes,
        differences,checked_items,idempotency_key
    ) values(
        p_assignment_id,v_assignment.order_id,v_assignment.assignee_email,
        v_assignment.mode,p_result,v_notes,p_differences,
        p_checked_items,p_idempotency_key
    ) returning id into v_id;

    update public.mobile_audit_assignments
       set status = case when p_result = 'No se pudo validar' then 'pendiente' else 'validado' end,
           last_result = p_result,
           updated_at = now()
     where id = p_assignment_id;

    return v_id::text;
end;
$$;

alter table public.mobile_audit_orders enable row level security;
alter table public.mobile_audit_assignments enable row level security;
alter table public.mobile_audit_reviews enable row level security;

-- No hay políticas RLS públicas. Sólo el servidor con clave privada
-- puede consultar y escribir; el navegador nunca recibe esa clave.
revoke all on public.mobile_audit_orders from public, anon, authenticated;
revoke all on public.mobile_audit_assignments from public, anon, authenticated;
revoke all on public.mobile_audit_reviews from public, anon, authenticated;
revoke execute on function public.submit_mobile_audit_review(uuid,text,text,text,jsonb,integer,text)
    from public, anon, authenticated;
revoke execute on function public.mobile_audit_check_assignment()
    from public, anon, authenticated;

grant select, insert on public.mobile_audit_orders to service_role;
grant select, insert on public.mobile_audit_assignments to service_role;
grant select, insert on public.mobile_audit_reviews to service_role;
grant update on public.mobile_audit_assignments to service_role;
grant execute on function public.submit_mobile_audit_review(uuid,text,text,text,jsonb,integer,text)
    to service_role;
