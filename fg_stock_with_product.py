from odoo import models, fields, api, _
from odoo.http import request

class FGStockDashboard(models.Model):
    _name = 'fg.stock.dashboard'
    _description = 'FG Stock Dashboard'

    @api.model
    def get_dashboard_data(self, filters):
        company_id = filters.get('company_id')
        salesperson_id = filters.get('salesperson_id')
        team_id = filters.get('team_id')
        
        if not company_id:
            company_id = self.env.company.id
        
        params = [company_id]
        
        # Base where clause common to all queries
        base_where = """
            WHERE od.next_operation = 'FG Packing'
              AND od.state NOT IN ('done', 'closed')
              AND od.fg_balance > 0
              AND od.company_id = %s
        """
        
        if salesperson_id:
            base_where += " AND so.user_id = %s "
            params.append(salesperson_id)
        if team_id:
            base_where += " AND so.team_id = %s "
            params.append(team_id)
            
        # 1. Payment Terms Aging
        query_payment_terms = f"""
            SELECT
                COALESCE(apt.name ->> 'en_US', apt.name::text) AS "Payment Terms",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_30+",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_30+",
                ROUND(SUM(od.fg_balance), 0) AS "Total_Balance",
                ROUND(SUM(od.fg_balance * od.price_unit), 0) AS "Total_Value"
            FROM operation_details od
            LEFT JOIN sale_order so ON od.oa_id = so.id
            LEFT JOIN account_payment_term apt ON so.payment_term_id = apt.id
            {base_where}
            GROUP BY COALESCE(apt.name ->> 'en_US', apt.name::text)
            ORDER BY "Payment Terms"
        """
        
        # 2. Item Category Aging
        query_item_category = f"""
            SELECT
                od.fg_categ_type AS "Item",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_30+",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_30+",
                ROUND(SUM(od.fg_balance), 0) AS "Total_Balance",
                ROUND(SUM(od.fg_balance * od.price_unit), 0) AS "Total_Value"
            FROM operation_details od
            LEFT JOIN sale_order so ON od.oa_id = so.id
            {base_where}
            GROUP BY od.fg_categ_type
            ORDER BY "Item"
        """
        
        # 3. Customer Aging
        query_customer = f"""
            SELECT
                rp.name AS "Customer",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_30+",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_30+",
                ROUND(SUM(od.fg_balance), 0) AS "Total_Balance",
                ROUND(SUM(od.fg_balance * od.price_unit), 0) AS "Total_Value"
            FROM operation_details od
            LEFT JOIN sale_order so ON od.oa_id = so.id
            LEFT JOIN res_partner rp ON so.partner_id = rp.id
            {base_where}
            GROUP BY rp.name
            ORDER BY "Customer"
        """
        
        # 4. Buyer Aging
        query_buyer = f"""
            SELECT
                od.buyer_name AS "Buyer",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance ELSE 0 END), 0) AS "Balance_30+",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_0-5",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_6-10",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_11-15",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_16-20",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_21-25",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_26-30",
                ROUND(SUM(CASE WHEN (CURRENT_DATE - od.action_date::date) > 30 THEN od.fg_balance * od.price_unit ELSE 0 END), 0) AS "Value_30+",
                ROUND(SUM(od.fg_balance), 0) AS "Total_Balance",
                ROUND(SUM(od.fg_balance * od.price_unit), 0) AS "Total_Value"
            FROM operation_details od
            LEFT JOIN sale_order so ON od.oa_id = so.id
            {base_where}
            GROUP BY od.buyer_name
            ORDER BY "Buyer"
        """
        
        self.env.cr.execute(query_payment_terms, tuple(params))
        payment_terms_data = self.env.cr.dictfetchall()
        
        self.env.cr.execute(query_item_category, tuple(params))
        item_category_data = self.env.cr.dictfetchall()
        
        self.env.cr.execute(query_customer, tuple(params))
        customer_data = self.env.cr.dictfetchall()
        
        self.env.cr.execute(query_buyer, tuple(params))
        buyer_data = self.env.cr.dictfetchall()
        
        return {
            'payment_terms': payment_terms_data,
            'item_category': item_category_data,
            'customer': customer_data,
            'buyer': buyer_data,
        }

    @api.model
    def get_drilldown_data(self, filters, group_type, group_value, bucket):
        company_id = filters.get('company_id')
        salesperson_id = filters.get('salesperson_id')
        team_id = filters.get('team_id')
        
        if not company_id:
            company_id = self.env.company.id
            
        params = [company_id]
        
        # Base where clause
        base_where = """
            WHERE od.next_operation = 'FG Packing'
              AND od.state NOT IN ('done', 'closed')
              AND od.fg_balance > 0
              AND od.company_id = %s
        """
        
        if salesperson_id:
            base_where += " AND so.user_id = %s "
            params.append(salesperson_id)
        if team_id:
            base_where += " AND so.team_id = %s "
            params.append(team_id)

        # Bucket condition
        bucket_map = {
            '0-5': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 0 AND 5",
            '6-10': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 6 AND 10",
            '11-15': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 11 AND 15",
            '16-20': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 16 AND 20",
            '21-25': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 21 AND 25",
            '26-30': "AND (CURRENT_DATE - od.action_date::date) BETWEEN 26 AND 30",
            '30+': "AND (CURRENT_DATE - od.action_date::date) > 30",
        }
        
        # Determine cleaned bucket name (remove 'Balance_' or 'Value_')
        clean_bucket = bucket.replace('Balance_', '').replace('Value_', '').replace('Total_', '')
        if clean_bucket in bucket_map:
            base_where += f" {bucket_map[clean_bucket]} "
        
        # Grouping filter
        if group_type == 'payment_terms':
            if group_value and group_value != 'Total':
                base_where += " AND COALESCE(apt.name ->> 'en_US', apt.name::text) = %s "
                params.append(group_value)
            elif group_value == 'Total':
                pass # Already handled by base_where
        elif group_type == 'item_category':
            if group_value and group_value != 'Total':
                base_where += " AND od.fg_categ_type = %s "
                params.append(group_value)
        elif group_type == 'customer':
            if group_value and group_value != 'Total':
                base_where += " AND rp.name = %s "
                params.append(group_value)
        elif group_type == 'buyer':
            if group_value and group_value != 'Total':
                base_where += " AND od.buyer_name = %s "
                params.append(group_value)

        query = f"""
            SELECT
                so.name AS "OA",
                od.fg_categ_type AS "Category",
                COALESCE(pt.name ->> 'en_US', pt.name::text) AS "Product",
                SUM(od.fg_balance) AS "Qty",
                SUM(ROUND(od.fg_balance * od.price_unit, 0)) AS "Value"
            FROM operation_details od
            LEFT JOIN sale_order so ON od.oa_id = so.id
            LEFT JOIN account_payment_term apt ON so.payment_term_id = apt.id
            LEFT JOIN res_partner rp ON so.partner_id = rp.id
            LEFT JOIN product_product pp ON od.product_id = pp.id
            LEFT JOIN product_template pt ON pp.product_tmpl_id = pt.id
            {base_where}
            GROUP BY so.name, od.fg_categ_type, pt.name
            ORDER BY so.name, pt.name
        """
        
        self.env.cr.execute(query, tuple(params))
        return self.env.cr.dictfetchall()

    @api.model
    def download_drilldown_excel(self, filters, group_type, group_value, bucket):
        data = self.get_drilldown_data(filters, group_type, group_value, bucket)
        
        import base64
        from io import BytesIO
        import xlsxwriter

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        sheet = workbook.add_worksheet("Drilldown Detail")

        # Formats
        header_format = workbook.add_format({'bold': True, 'bg_color': '#8a2b5d', 'font_color': 'white', 'border': 1, 'align': 'center'})
        num_format = workbook.add_format({'num_format': '#,##0', 'border': 1})
        text_format = workbook.add_format({'border': 1})
        total_format = workbook.add_format({'bold': True, 'bg_color': '#f1f5f9', 'border': 1, 'num_format': '#,##0'})

        # Headers
        headers = ["OA", "Category", "Product", "Qty", "Value"]
        for col, header in enumerate(headers):
            sheet.write(0, col, header, header_format)
            sheet.set_column(col, col, 15 if col < 3 else 12)

        # Data
        row_idx = 1
        total_qty = 0
        total_value = 0
        for item in data:
            sheet.write(row_idx, 0, item.get('OA', ''), text_format)
            sheet.write(row_idx, 1, item.get('Category', ''), text_format)
            sheet.write(row_idx, 2, item.get('Product', ''), text_format)
            sheet.write(row_idx, 3, item.get('Qty', 0), num_format)
            sheet.write(row_idx, 4, item.get('Value', 0), num_format)
            total_qty += item.get('Qty', 0)
            total_value += item.get('Value', 0)
            row_idx += 1

        # Totals
        sheet.write(row_idx, 0, "Total", total_format)
        sheet.write(row_idx, 1, "", total_format)
        sheet.write(row_idx, 2, "", total_format)
        sheet.write(row_idx, 3, total_qty, total_format)
        sheet.write(row_idx, 4, total_value, total_format)

        workbook.close()
        output.seek(0)
        
        filename = f"{group_value}_{bucket}.xlsx".replace(" ", "_")
        return {
            'file_data': base64.b64encode(output.read()).decode(),
            'filename': filename
        }

    @api.model
    def download_excel(self, filters):
        import base64
        from io import BytesIO
        from datetime import datetime
        try:
            import xlsxwriter
        except ImportError:
            raise ImportError("Please install xlsxwriter: pip install xlsxwriter")

        data = self.get_dashboard_data(filters)
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Styling
        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4F46E5', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        sub_header_format = workbook.add_format({
            'bold': True, 'bg_color': '#F1F5F9', 'align': 'center', 'valign': 'vcenter', 'border': 1
        })
        data_format = workbook.add_format({'border': 1})
        num_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
        total_format = workbook.add_format({'bold': True, 'bg_color': '#E0E7FF', 'border': 1})
        total_num_format = workbook.add_format({'bold': True, 'bg_color': '#E0E7FF', 'border': 1, 'num_format': '#,##0'})

        # Summary Sheet
        summary_sheet = workbook.add_worksheet('Summary')
        summary_sheet.set_column(0, 0, 25)
        summary_sheet.set_column(1, 1, 15)

        # Calculate Totals for Summary
        total_balance = sum(row.get('Total_Balance', 0) for row in data['payment_terms'])
        total_value = sum(row.get('Total_Value', 0) for row in data['payment_terms'])
        total05_bal = sum(row.get('Balance_0-5', 0) for row in data['payment_terms'])
        total05_val = sum(row.get('Value_0-5', 0) for row in data['payment_terms'])
        total30_plus_bal = sum(row.get('Balance_30+', 0) for row in data['payment_terms'])
        total30_plus_val = sum(row.get('Value_30+', 0) for row in data['payment_terms'])

        summary_rows = [
            ['Metric', 'Value'],
            ['Total FG Balance', total_balance],
            ['Total FG Value', total_value],
            ['', ''],
            ['0-5 Days Summary', ''],
            ['Balance', total05_bal],
            ['Value', total05_val],
            ['', ''],
            ['30+ Days Summary', ''],
            ['Balance', total30_plus_bal],
            ['Value', total30_plus_val],
        ]

        for r, row in enumerate(summary_rows):
            if r == 0:
                for c, val in enumerate(row):
                    summary_sheet.write(r, c, val, header_format)
            elif r in [4, 8]:
                summary_sheet.write(r, 0, row[0], sub_header_format)
                summary_sheet.write(r, 1, '', sub_header_format)
            elif row[0]:
                summary_sheet.write(r, 0, row[0], data_format)
                summary_sheet.write(r, 1, row[1], num_format)

        # Tables (Separate Sheets)
        def write_aging_table(sheet_name, title, label, rows):
            sheet = workbook.add_worksheet(sheet_name)
            sheet.set_column(0, 0, 25)
            sheet.set_column(1, 16, 12)
            
            sheet.merge_range(0, 0, 0, 16, title, header_format)
            
            headers_top = [
                label, '0-5 Days', '6-10 Days', '11-15 Days', '16-20 Days', 
                '21-25 Days', '26-30 Days', '30+ Days', 'Total'
            ]
            
            # Write Top Headers (Merged)
            sheet.merge_range(1, 0, 2, 0, label, sub_header_format)
            col = 1
            for h in headers_top[1:-1]:
                sheet.merge_range(1, col, 1, col + 1, h, sub_header_format)
                sheet.write(2, col, 'Balance', sub_header_format)
                sheet.write(2, col + 1, 'Value', sub_header_format)
                col += 2
            
            sheet.merge_range(1, 15, 1, 16, 'Total', sub_header_format)
            sheet.write(2, 15, 'Total Balance', sub_header_format)
            sheet.write(2, 16, 'Total Value', sub_header_format)

            # Calculate and Write Grand Totals for this table
            cols_to_sum = [
                'Balance_0-5', 'Value_0-5', 'Balance_6-10', 'Value_6-10',
                'Balance_11-15', 'Value_11-15', 'Balance_16-20', 'Value_16-20',
                'Balance_21-25', 'Value_21-25', 'Balance_26-30', 'Value_26-30',
                'Balance_30+', 'Value_30+', 'Total_Balance', 'Total_Value'
            ]
            
            t_row = 3
            sheet.write(t_row, 0, 'Total', total_format)
            for i, col_key in enumerate(cols_to_sum):
                val = sum(r.get(col_key, 0) for r in rows)
                sheet.write(t_row, i + 1, val, total_num_format)

            # Write Data Rows
            current_r = t_row + 1
            for r_data in rows:
                sheet.write(current_r, 0, r_data.get(label, ''), data_format)
                for i, col_key in enumerate(cols_to_sum):
                    val = r_data.get(col_key, 0)
                    sheet.write(current_r, i + 1, val, num_format)
                current_r += 1

        write_aging_table('Payment Terms', 'Aging by Payment Terms', 'Payment Terms', data['payment_terms'])
        write_aging_table('Item Category', 'Aging by Item Category', 'Item', data['item_category'])
        write_aging_table('Customer', 'Aging by Customer', 'Customer', data['customer'])
        write_aging_table('Buyer', 'Aging by Buyer', 'Buyer', data['buyer'])

        workbook.close()
        output.seek(0)
        
        excel_data = base64.b64encode(output.read()).decode('utf-8')
        filename = f"FG_Stock_Aging_{datetime.now().strftime('%Y%m%d')}.xlsx"
        
        return {
            'file_data': excel_data,
            'filename': filename
        }
