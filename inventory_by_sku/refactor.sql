with stock_on_hand AS (
    SELECT
        dim_calendar.date,
        dim_sites.combo,
        dim_products.id AS product_id,
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
        dim_products.type != 'Point of sales'
        AND
        dim_calendar.date = now()::date
)

SELECT
    soh.combo,
    soh.product_id,
	sum(soh.soh_standard_value ) AS soh_standard_value,
    sum(soh.soh_converted_quantity) AS soh_converted_quantity,
    sum((soh.soh_converted_quantity * product.weight) / 1000) AS soh_volume
FROM
    stock_on_hand AS soh
JOIN 
    dim_products as product 
ON 
    product_id = product.id                    
    ,
    (
        SELECT
             now()::date AS date
    ) AS t
WHERE
    1 = 1
    AND soh.status IN ('Active', 'Discontinued', 'To be discontinued')
    AND soh.type IN ('New', 'Promotion', 'Standard')
    AND soh.date = t.date
GROUP BY 1, 2
