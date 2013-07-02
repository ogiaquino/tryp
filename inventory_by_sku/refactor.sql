with 
stock_on_hand AS (
    SELECT
        dim_products.id AS product_id,
        dim_calendar.date,
        dim_sites.combo,
        dim_products.status,
        dim_products.type,
        (coalesce(stocks.quantity_on_hand, 0)::float / dim_products.conversion_factor)::numeric AS soh_converted_quantity,
        (coalesce(stocks.standard_value, 0)) AS soh_standard_value,
        (coalesce(stocks.quantity_on_hand, 0)::float / dim_products.conversion_factor)::numeric * dim_products.weight AS soh_volume
    FROM
        mv_fact_daily_good_stocks as stocks
    JOIN 
        dim_products ON product_id = dim_products.id                          
    JOIN 
        dim_sites ON site_id = dim_sites.id
    JOIN 
        dim_calendar ON calendar_id = dim_calendar.id
    WHERE                                                                      
        dim_calendar.date = now()::date
        AND dim_products.status IN ('Active', 'Discontinued', 'To be discontinued')
        AND dim_products.type IN ('New', 'Promotion', 'Standard')
),

sales AS (
    SELECT
        dim_products.id AS product_id,
        dim_calendar.date,
        dim_sites.combo,
        dim_products.status,
        dim_products.type,
        (coalesce(invoices.delivered_quantity, 0)::float / dim_products.conversion_factor)::numeric AS sales_converted_quantity,
        coalesce(invoices.standard_value, 0) AS sales_standard_value,
        (coalesce(invoices.delivered_quantity, 0)::float / dim_products.conversion_factor)::numeric * dim_products.weight AS sales_volume
    FROM
        mv_fact_invoices as invoices
    JOIN 
        dim_sites ON site_id = dim_sites.id 
    JOIN 
        dim_products ON product_id = dim_products.id                          
    JOIN 
        dim_calendar ON calendar_id = dim_calendar.id
    WHERE                                                                      
        dim_calendar.date BETWEEN (now()::date - '1 day'::interval - '74 day'::interval) AND (now()::date - '1 day'::interval)
        AND dim_products.status IN ('Active', 'Discontinued', 'To be discontinued')
        AND dim_products.type IN ('New', 'Promotion', 'Standard')
)

SELECT * FROM sales FULL JOIN stock_on_hand using(product_id,date,combo,status,type) LIMIT 5;

    SELECT
        dim_products.id AS product_id,
        dim_calendar.date,
        dim_sites.combo,
        dim_products.status,
        dim_products.type,
        (coalesce(invoices.delivered_quantity, 0)::float / dim_products.conversion_factor)::numeric AS sales_converted_quantity,
        coalesce(invoices.standard_value, 0) AS sales_standard_value,
        (coalesce(invoices.delivered_quantity, 0)::float / dim_products.conversion_factor)::numeric * dim_products.weight AS sales_volume
    FROM
        mv_fact_invoices as invoices
    JOIN 
        dim_sites ON site_id = dim_sites.id 
    JOIN 
        dim_products ON product_id = dim_products.id                          
    JOIN 
        dim_calendar ON calendar_id = dim_calendar.id
    WHERE                                                                      
        dim_calendar.date BETWEEN (now()::date - '1 day'::interval - '74 day'::interval) AND (now()::date - '1 day'::interval)
        AND dim_products.status IN ('Active', 'Discontinued', 'To be discontinued')
        AND dim_products.type IN ('New', 'Promotion', 'Standard')
