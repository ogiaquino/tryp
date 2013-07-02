with
stock_on_hand AS (
    SELECT
        dim_products.id AS product_id,
        dim_calendar.date,
        dim_sites.combo,
        dim_products.status,
        dim_products.type,
        (coalesce(stocks.quantity_on_hand, 0)::float /
            dim_products.conversion_factor)::numeric AS soh_converted_quantity,
        (coalesce(stocks.standard_value, 0)) AS soh_standard_value,
        (coalesce(stocks.quantity_on_hand, 0)::float /
            dim_products.conversion_factor)::numeric * dim_products.weight AS soh_volume
    FROM
        mv_fact_daily_good_stocks as stocks
    JOIN dim_products
        ON product_id = dim_products.id
    JOIN dim_sites
        ON site_id = dim_sites.id
    JOIN dim_calendar
        ON calendar_id = dim_calendar.id
    WHERE
        dim_calendar.date = now()::date
        --AND dim_products.type != 'Point of sales'
        AND dim_products.status IN ('Active', 'Discontinued', 'To be discontinued')
        AND dim_products.type IN ('New', 'Promotion', 'Standard')
        AND dim_sites.principal_code = 'kraft'
        --AND dim_sites.dome_identifier || '.' || dim_sites.principal_code IN (%(security_filter)s)
        AND dim_sites.active IS TRUE
),

sales AS (
    SELECT
        dim_products.id AS product_id,
        dim_calendar.date,
        dim_sites.combo,
        dim_products.status,
        dim_products.type,
        (coalesce(invoices.delivered_quantity, 0)::float /
            dim_products.conversion_factor)::numeric AS sales_converted_quantity,
        coalesce(invoices.standard_value, 0) AS sales_standard_value,
        (coalesce(invoices.delivered_quantity, 0)::float /
            dim_products.conversion_factor)::numeric * dim_products.weight AS sales_volume
    FROM
        mv_fact_invoices as invoices
    JOIN dim_sites
        ON site_id = dim_sites.id
    JOIN dim_products
        ON product_id = dim_products.id
    JOIN dim_calendar
        ON calendar_id = dim_calendar.id
    WHERE
        dim_calendar.date BETWEEN (now()::date - '1 day'::interval - '74 day'::interval)
            AND (now()::date - '1 day'::interval)
        --AND dim_products.type != 'Point of sales'
        AND dim_products.status IN ('Active', 'Discontinued', 'To be discontinued')
        AND dim_products.type IN ('New', 'Promotion', 'Standard')
        AND dim_sites.principal_code = 'kraft'
        --AND dim_sites.dome_identifier || '.' || dim_sites.principal_code IN (%(security_filter)s)
        AND dim_sites.active IS TRUE
        AND invoices.delivered_quantity > 0
)

SELECT
    sum(soh_value),
    sum("Sales Value")
FROM
(
SELECT
    "Category",
    "Brand",
    "SKU",
    "Description",
    "Region Parent",
    "Region",
    "Distributor",
    sum(sales_standard_value) AS "Sales Value",
    sum(sales_converted_quantity) AS "Sales Qty",
    sum(sales_volume) AS "Sales Volume",
    sum(soh_standard_value) AS soh_value,
    sum(soh_converted_quantity) AS "SOH Qty",
    sum(soh_volume) AS "SOH Volume"
FROM
(
    SELECT
        categories.parent_name AS "Category",
        brands.parent_name AS "Brand",
        brands.child_code AS "SKU",
        brands.child_name AS "Description",
        regions.parent_name AS "Region Parent",
        areas.parent_name AS "Region",
        distributors.parent_name AS "Distributor",
        sales_standard_value,
        sales_converted_quantity,
        sales_volume,
        soh_standard_value,
        soh_converted_quantity,
        soh_volume
    FROM
        sales
    FULL JOIN stock_on_hand
    using(product_id,combo)
    JOIN dim_sites
    USING (combo)
    JOIN mv_closure_sites_hierarchy AS areas
    ON
        dim_sites.principal_code = areas.principal_code
        AND dim_sites.dome_identifier = areas.dome_identifier
        AND areas.extended IS FALSE
        AND areas.parent_level = 3
    JOIN mv_closure_sites_hierarchy AS distributors
    ON
        dim_sites.principal_code = distributors.principal_code
        AND dim_sites.dome_identifier = distributors.dome_identifier
        AND distributors.extended IS FALSE
        AND distributors.parent_level = 4
    JOIN mv_closure_sites_hierarchy AS regions
    ON
        dim_sites.principal_code = regions.principal_code
        AND dim_sites.dome_identifier = regions.dome_identifier
        AND regions.extended IS FALSE
        AND regions.parent_level = 2
    JOIN mv_closure_product_categories AS categories
    ON
        product_id = categories.child_id
        AND categories.extended IS TRUE
        AND categories.parent_level = 1
    JOIN mv_closure_product_categories AS brands
    ON
        product_id = brands.child_id
        AND brands.extended IS TRUE
        AND brands.parent_level = 2

) as FOO GROUP BY 1,2,3,4,5,6,7
) as TET
