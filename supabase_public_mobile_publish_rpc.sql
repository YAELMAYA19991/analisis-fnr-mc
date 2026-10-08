-- Let the public mobile page publish only audit copies and create an assignment.
-- These RPCs never update or delete source orders or application data.
create or replace function public.publish_mobile_audit_order_public(
    p_order jsonb,
    p_mode text,
    p_assignee_email text
) returns text
language plpgsql
security definer
set search_path = public, pg_temp
as $function$
declare
    v_owners text[];
    v_assignment_id uuid;
    v_assignee text := lower(btrim(coalesce(p_assignee_email, '')));
    v_id text := coalesce(p_order->>'id', '');
    v_batch text := coalesce(p_order->>'batch_id', '');
    v_items jsonb := p_order->'items';
begin
    if jsonb_typeof(p_order) <> 'object' then
        raise exception 'El pedido no tiene un formato válido.';
    end if;
    if v_id !~ '^[0-9a-f]{64}$' or v_batch !~ '^[0-9a-f]{64}$' then
        raise exception 'El identificador del pedido no es válido.';
    end if;
    if nullif(btrim(coalesce(p_order->>'pedido','')), '') is null
       or length(p_order->>'pedido') > 100 then
        raise exception 'El número de pedido está vacío o es demasiado largo.';
    end if;
    if jsonb_typeof(v_items) <> 'array'
       or jsonb_array_length(v_items) < 1
       or jsonb_array_length(v_items) > 2000 then
        raise exception 'El pedido debe contener entre 1 y 2000 artículos.';
    end if;
    if jsonb_typeof(p_order->'owner_emails') <> 'array'
       or jsonb_array_length(p_order->'owner_emails') < 1 then
        raise exception 'No se identificó al menos un picker del pedido.';
    end if;

    select array_agg(distinct lower(btrim(value)))
      into v_owners
      from jsonb_array_elements_text(p_order->'owner_emails') as value;

    if cardinality(v_owners) < 1
       or exists (
            select 1 from unnest(v_owners) e
            where e !~ '^[a-z0-9._%+\-]+@justo[.]mx$'
       ) then
        raise exception 'El picker debe tener correo corporativo @justo.mx.';
    end if;
    if exists (
        select 1
        from jsonb_array_elements(v_items) as item(value)
        where nullif(btrim(item.value->>'picker_email'), '') is not null
          and not (lower(item.value->>'picker_email') = any(v_owners))
    ) then
        raise exception 'El pedido contiene un picker que no está en la plantilla.';
    end if;

    if p_mode = 'autoverificacion' then
        if not (v_assignee = any(v_owners))
           or not exists (
                select 1 from jsonb_array_elements(v_items) as item(value)
                where lower(coalesce(item.value->>'picker_email','')) = v_assignee
           ) then
            raise exception 'La autoverificación requiere un picker identificado del pedido.';
        end if;
    elsif p_mode = 'auditoria_cruzada' then
        if v_assignee <> 'auditoria-publica@coyoacan.invalid' then
            raise exception 'La auditoría cruzada debe asignarse a la cola pública de supervisión.';
        end if;
    else
        raise exception 'Tipo de revisión no permitido.';
    end if;

    insert into public.mobile_audit_orders(
        id,batch_id,pedido,slot,owner_emails,picker_labels,items,published_by_email
    ) values (
        v_id,
        v_batch,
        btrim(p_order->>'pedido'),
        coalesce(p_order->>'slot',''),
        v_owners,
        coalesce(p_order->>'picker_labels',''),
        v_items,
        'acceso-publico'
    )
    on conflict (id) do nothing;

    insert into public.mobile_audit_assignments(
        order_id,mode,assignee_email,assigned_by_email
    ) values (
        v_id,p_mode,v_assignee,'acceso-publico'
    )
    on conflict (order_id,mode,assignee_email) do nothing
    returning id into v_assignment_id;

    if v_assignment_id is null then
        select id into v_assignment_id
          from public.mobile_audit_assignments
         where order_id=v_id and mode=p_mode and assignee_email=v_assignee;
    end if;

    return v_assignment_id::text;
end;
$function$;

create or replace function public.list_mobile_audit_reviews_public(p_order_id text)
returns table (
    reviewer_name text,
    mode text,
    result text,
    notes text,
    differences jsonb,
    created_at timestamptz
)
language sql
security definer
set search_path = public, pg_temp
stable
as $function$
    select r.reviewer_name, r.mode, r.result, r.notes, r.differences, r.created_at
    from public.mobile_audit_reviews r
    where r.order_id = p_order_id
    order by r.created_at desc
    limit 100;
$function$;

revoke all on function public.publish_mobile_audit_order_public(jsonb,text,text) from public, anon, authenticated;
grant execute on function public.publish_mobile_audit_order_public(jsonb,text,text) to anon;

revoke all on function public.list_mobile_audit_reviews_public(text) from public, anon, authenticated;
grant execute on function public.list_mobile_audit_reviews_public(text) to anon;

revoke all on table public.mobile_audit_orders,
    public.mobile_audit_assignments,
    public.mobile_audit_reviews
    from public, anon;

notify pgrst, 'reload schema';
