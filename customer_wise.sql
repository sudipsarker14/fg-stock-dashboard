SELECT
	rp.name AS "Customer",
    
	-- FG Balance by Aging Bucket
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance ELSE 0 END) AS "Balance_0-5",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance ELSE 0 END) AS "Balance_6-10",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance ELSE 0 END) AS "Balance_11-15",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance ELSE 0 END) AS "Balance_16-20",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance ELSE 0 END) AS "Balance_21-25",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance ELSE 0 END) AS "Balance_26-30",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance ELSE 0 END) AS "Balance_30+",

    -- Value by Aging Bucket
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_0-5",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_6-10",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_11-15",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_16-20",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_21-25",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_26-30",
    SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance * od.price_unit ELSE 0 END) AS "Value_30+",

    -- Grand totals
    SUM(od.fg_balance) AS "Total_Balance",
    SUM(od.fg_balance * od.price_unit) AS "Total_Value"

FROM operation_details od

LEFT JOIN sale_order so
       ON od.oa_id = so.id
LEFT JOIN res_partner rp ON so.partner_id = rp.id

WHERE
      od.next_operation = 'FG Packing'
  AND od.state NOT IN ('done', 'closed')
  AND od.fg_balance > 0

GROUP BY rp.name
ORDER BY "Customer";
