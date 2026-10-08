-- Public mobile audit access, narrowed to sanitized RPC results.
-- The orders, assignments, and review tables remain inaccessible directly to anon.
-- No existing rows or original orders are changed.
create or replace function public.list_mobile_audit_assignments_public()
returns table (
    id uuid,
    mode text,
    status text,
    last_result text,
    order_id text,
    mobile_audit_orders jsonb
)
language sql
security definer
set search_path = public, pg_temp
stable
as $function$
    select
        a.id,
        a.mode,
        a.status,
        a.last_result,
        a.order_id,
        jsonb_build_object(
            'pedido', o.pedido,
            'slot', o.slot,
            'picker_labels', o.picker_labels,
            'items', safe_items.items
        ) as mobile_audit_orders
    from public.mobile_audit_assignments as a
    join public.mobile_audit_orders as o on o.id = a.order_id
    cross join lateral (
        select coalesce(
            jsonb_agg(
                jsonb_build_object(
                    'sku', item.value->>'sku',
                    'articulo', item.value->>'articulo',
                    'cantidad_pedida', item.value->>'cantidad_pedida',
                    'cantidad_pickeada', item.value->>'cantidad_pickeada'
                ) order by item.ordinality
            ),
            '[]'::jsonb
        ) as items
        from jsonb_array_elements(o.items) with ordinality as item(value, ordinality)
        where a.mode = 'auditoria_cruzada'
           or item.value->>'picker_email' = a.assignee_email
    ) as safe_items
    order by a.created_at desc
    limit 500;
$function$;

revoke all on function public.list_mobile_audit_assignments_public() from public, anon, authenticated;
grant execute on function public.list_mobile_audit_assignments_public() to anon;

revoke all on function public.submit_mobile_audit_review_supervisor(uuid,text,text,text,jsonb,integer,text)
    from public, anon, authenticated;
grant execute on function public.submit_mobile_audit_review_supervisor(uuid,text,text,text,jsonb,integer,text)
    to anon;

revoke all on table public.mobile_audit_orders,
    public.mobile_audit_assignments,
    public.mobile_audit_reviews
    from public, anon;

notify pgrst, 'reload schema';