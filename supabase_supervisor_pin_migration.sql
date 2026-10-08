-- Migración aditiva para auditoría móvil con código privado de supervisor.
-- No elimina ni reescribe registros existentes.
alter table public.mobile_audit_reviews
    add column if not exists reviewer_name text not null default '';

create or replace function public.submit_mobile_audit_review_supervisor(
    p_assignment_id uuid,
    p_reviewer_name text,
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
    v_assignee_email text;
    v_review_id text;
    v_reviewer_name text := btrim(coalesce(p_reviewer_name, ''));
begin
    if v_reviewer_name = '' or length(v_reviewer_name) > 100 then
        raise exception 'Escribe el nombre del supervisor (máximo 100 caracteres).';
    end if;

    select assignee_email into v_assignee_email
      from public.mobile_audit_assignments
     where id = p_assignment_id;

    if not found then
        raise exception 'Asignación inexistente.';
    end if;

    v_review_id := public.submit_mobile_audit_review(
        p_assignment_id, v_assignee_email, p_result, p_notes,
        p_differences, p_checked_items, p_idempotency_key
    );

    update public.mobile_audit_reviews
       set reviewer_name = v_reviewer_name
     where id = v_review_id::uuid;

    return v_review_id;
end;
$$;

revoke all on function public.submit_mobile_audit_review_supervisor(uuid,text,text,text,jsonb,integer,text)
    from public, anon, authenticated;
grant execute on function public.submit_mobile_audit_review_supervisor(uuid,text,text,text,jsonb,integer,text)
    to service_role;