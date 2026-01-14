# -*- coding: utf-8 -*-
#############################################################################
#
#    Cybrosys Technologies Pvt. Ltd.
#
#    Copyright (C) 2024-TODAY Cybrosys Technologies(<https://www.cybrosys.com>)
#    Author: Cybrosys Techno Solutions(<https://www.cybrosys.com>)
#
#    You can modify it under the terms of the GNU LESSER
#    GENERAL PUBLIC LICENSE (AGPL v3), Version 3.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU LESSER GENERAL PUBLIC LICENSE (AGPL v3) for more details.
#
#    You should have received a copy of the GNU LESSER GENERAL PUBLIC LICENSE
#    (AGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################
import io
import xlsxwriter
from odoo import fields, models
from odoo.tools import date_utils
import json
from datetime import datetime, time


class TurnoverReport(models.TransientModel):
    """Wizard created for to select date, products, categories, companies and
    warehouses. The records are filtered by using this fields"""

    _name = "turnover.report"
    _description = "Turnover Report"

    start_date = fields.Date(string="Start Date", help="Select inventory start date.")
    end_date = fields.Date(string="End Date", help="Select inventory end date.")
    product_ids = fields.Many2many(
        "product.product",
        string="Products",
        default=lambda self: self.env["product.product"].search([], limit=1),
        help="Select multiple products " "from the list.",
    )
    category_ids = fields.Many2many(
        "product.category",
        string="Category",
        default=lambda self: self._default_categ_ids(),
        help="Select multiple categories " "from the list",
    )
    warehouse_ids = fields.Many2many(
        "stock.warehouse",
        string="Warehouse",
        default=lambda self: self._default_warehouse_ids(),
        help="Select multiple warehouses " "from the list.",
    )
    company_ids = fields.Many2many(
        "res.company",
        string="Company",
        default=lambda self: self.env.company,
        help="Select multiple companies " "from the list.",
    )

    def _default_categ_ids(self):
        """Return default category to selection field."""
        category = self.env["product.product"].search([], limit=1).product_tmpl_id.categ_id
        return [(6, 0, [category.id])] if category else []

    def _default_warehouse_ids(self):
        """Return default warehouse to selection field."""
        warehouse = self.env["stock.warehouse"].search([], limit=1)
        return [(6, 0, [warehouse.id])] if warehouse else []

    def action_pdf_report_generate(self):
        """Here generate a dictionary of list of datas and that return to a
        report action. And it will generate the pdf report."""
        data = {
            "stock_report": self.call_render_report(),
            "start_date": self.start_date,
            "end_date": self.end_date,
        }
        return self.env.ref("inventory_turnover_report_analysis.inventory_turnover_report").report_action(
            self, data=data
        )

    def action_xlsx_report_generate(self):
        """Here generate a dictionary of list of datas and that return to a
        report action. And it will generate the xlsx report."""
        data = {
            "stock_report": self.call_render_report(),
            "start_date": self.start_date,
            "end_date": self.end_date,
        }
        return {
            "type": "ir.actions.report",
            "data": {
                "model": "turnover.report",
                "options": json.dumps(data, default=date_utils.json_default),
                "output_format": "xlsx",
                "report_name": "Inventory Turnover Analysis Report",
            },
            "report_type": "xlsx",
        }

    def get_xlsx_report(self, data, response):
        """This function is for create xlsx report"""
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Inventory Turnover Analysis Report")
        head = workbook.add_format({"align": "center", "bold": True, "font_size": "30px"})
        sheet.set_column("A:A", 30)
        sheet.set_column("B:B", 15)
        sheet.set_column("C:C", 15)
        sheet.set_column("D:D", 15)
        sheet.set_column("E:E", 15)
        sheet.set_column("F:F", 15)
        sheet.set_column("G:G", 15)
        sheet.merge_range("A3:G1", "Inventory Turnover Analysis Report", head)
        row = 6
        column = 0
        if data["start_date"]:
            sheet.write(5, 1, "Start Date:", workbook.add_format({"align": "center", "bold": True}))
            sheet.write(5, 2, data["start_date"], workbook.add_format({"align": "center", "bold": True}))
            row += 1
        if data["end_date"]:
            sheet.write(5, 4, "End Date:", workbook.add_format({"align": "center", "bold": True}))
            sheet.write(5, 5, data["end_date"], workbook.add_format({"align": "center", "bold": True}))
            row += 1
        head_table = workbook.add_format({"align": "center", "bold": True})
        sheet.write(row, column, "Product", workbook.add_format({"align": "left", "bold": True}))
        column += 1
        sheet.write(row, column, "Opening Stock", head_table)
        column += 1
        sheet.write(row, column, "Closing Stock", head_table)
        column += 1
        sheet.write(row, column, "Average Stock", head_table)
        column += 1
        sheet.write(row, column, "Sale count", head_table)
        column += 1
        sheet.write(row, column, "Purchase Count", head_table)
        column += 1
        sheet.write(row, column, "Turnover Ratio", head_table)
        for datas in data["stock_report"]:
            row += 1
            column = 0
            table_body = workbook.add_format({"align": "center"})
            sheet.write(row, column, datas["product"], workbook.add_format({"align": "left"}))
            column += 1
            sheet.write(row, column, datas["opening_stock"], table_body)
            column += 1
            sheet.write(row, column, datas["closing_stock"], table_body)
            column += 1
            sheet.write(row, column, datas["average_stock"], table_body)
            column += 1
            sheet.write(row, column, datas["sale_count"], table_body)
            column += 1
            sheet.write(row, column, datas["purchase_count"], table_body)
            column += 1
            sheet.write(row, column, datas["turnover_ratio"], table_body)
        workbook.close()
        output.seek(0)
        response.stream.write(output.read())
        output.close()

    def action_data_fetch(self):
        """Here generate a list of dictionary of datas, and from that list
        create records. And it will return tree view with created records."""
        self.env["fetch.data"].search([]).unlink()
        filtered_records = self.call_render_model()
        for rec in filtered_records:
            self.env["fetch.data"].create(
                {
                    "company_id": rec["company_id"],
                    "warehouse_id": rec["warehouse_id"],
                    "product_id": rec["id"],
                    "category_id": rec["category_id"],
                    "opening_stock": rec["opening_stock"],
                    "closing_stock": rec["closing_stock"],
                    "average_stock": rec["average_stock"],
                    "sale_count": rec["sale_count"],
                    "purchase_count": rec["purchase_count"],
                    "turnover_ratio": rec["turnover_ratio"],
                }
            )
        return {
            "type": "ir.actions.act_window",
            "view_mode": "tree",
            "res_model": "fetch.data",
            "name": "Turnover Analysis Report",
            "target": "current",
            "context": {"create": False},
        }

    def action_generate_graph_view(self):
        """Here generate a list of dictionary of datas, and from that list
        create records. And it will return graph view with created records."""
        self.env["turnover.graph.analysis"].search([]).unlink()
        filtered_records = self.call_render_model()
        for rec in filtered_records:
            self.env["turnover.graph.analysis"].create(
                {
                    "company_id": rec["company_id"],
                    "warehouse_id": rec["warehouse_id"],
                    "product_id": rec["id"],
                    "category_id": rec["category_id"],
                    "opening_stock": rec["opening_stock"],
                    "closing_stock": rec["closing_stock"],
                    "average_stock": rec["average_stock"],
                    "sale_count": rec["sale_count"],
                    "purchase_count": rec["purchase_count"],
                    "turnover_ratio": rec["turnover_ratio"],
                }
            )
        return {
            "type": "ir.actions.act_window",
            "view_mode": "graph",
            "res_model": "turnover.graph.analysis",
            "name": "Turnover Analysis",
            "target": "current",
            "context": {"create": False},
        }

    def call_render_report(self):
        """Function call for get datas to generate PDF and XLSX report, and
        return the computed record."""
        return self._compute_turnover_data(for_report=True)

    def call_render_model(self):
        """Function call for get datas to generate list and graph view, and
        return the computed record."""
        return self._compute_turnover_data(for_report=False)

    def _compute_turnover_data(self, for_report=True):
        """Compute inventory turnover data based on stock moves within date range.
        OPTIMIZED: Uses batch queries instead of per-product/location loops

        Args:
            for_report: If True, returns data formatted for reports (with name strings),
                       If False, returns data formatted for models (with IDs)
        """
        stock_report = []

        # Build product domain
        product_domain = []
        if self.product_ids:
            product_domain.append(("id", "in", self.product_ids.ids))
        if self.category_ids:
            product_domain.append(("categ_id", "in", self.category_ids.ids))

        products = self.env["product.product"].search(product_domain)
        if not products:
            return stock_report

        # Build location domain for warehouses
        location_ids = []
        if self.warehouse_ids:
            for warehouse in self.warehouse_ids:
                location_ids += (
                    self.env["stock.location"]
                    .search([("location_id", "child_of", warehouse.view_location_id.id), ("usage", "=", "internal")])
                    .ids
                )
        else:
            location_ids = self.env["stock.location"].search([("usage", "=", "internal")]).ids

        if not location_ids:
            return stock_report

        # Get warehouse/company mapping
        location_warehouse_map = {}
        for location in self.env["stock.location"].browse(location_ids):
            location_warehouse_map[location.id] = {
                "warehouse": location.warehouse_id,
                "company": location.company_id or self.env.company,
            }

        # BATCH QUERY: Get all stock data at once using SQL
        opening_stock_data = self._get_batch_stock_at_date(products.ids, location_ids, self.start_date)
        closing_stock_data = self._get_batch_stock_at_date(products.ids, location_ids, self.end_date)
        sales_data = self._get_batch_sales(products.ids, location_ids)
        purchase_data = self._get_batch_purchases(products.ids, location_ids)

        # Aggregate by product, warehouse, company
        aggregated_data = {}

        for product in products:
            for location_id in location_ids:
                loc_info = location_warehouse_map[location_id]
                warehouse = loc_info["warehouse"]
                company = loc_info["company"]

                # Filter by selected companies
                if self.company_ids and company.id not in self.company_ids.ids:
                    continue

                key = (product.id, warehouse.id, company.id)

                if key not in aggregated_data:
                    aggregated_data[key] = {
                        "product": product,
                        "warehouse": warehouse,
                        "company": company,
                        "opening_stock": 0,
                        "closing_stock": 0,
                        "sale_count": 0,
                        "purchase_count": 0,
                    }

                # Add stock data
                opening_key = (product.id, location_id)
                closing_key = (product.id, location_id)
                sales_key = (product.id, location_id)
                purchase_key = (product.id, location_id)

                aggregated_data[key]["opening_stock"] += opening_stock_data.get(opening_key, 0)
                aggregated_data[key]["closing_stock"] += closing_stock_data.get(closing_key, 0)
                aggregated_data[key]["sale_count"] += sales_data.get(sales_key, 0)
                aggregated_data[key]["purchase_count"] += purchase_data.get(purchase_key, 0)

        # Format results
        for (product_id, warehouse_id, company_id), data in aggregated_data.items():
            product = data["product"]
            opening_stock = data["opening_stock"]
            closing_stock = data["closing_stock"]
            sale_count = data["sale_count"]
            purchase_count = data["purchase_count"]

            # Calculate average stock
            average_stock = (opening_stock + closing_stock) / 2

            # Calculate turnover ratio: COGS / Average Stock
            turnover_ratio = 0
            if average_stock > 0 and sale_count > 0:
                turnover_ratio = sale_count / average_stock

            turnover = round(turnover_ratio, 2)

            # Format product name
            name = product.display_name
            split_name = name.split("]")
            product_name = split_name[1].strip() if len(split_name) > 1 else split_name[0]

            if for_report:
                values = {
                    "id": product.id,
                    "product": product_name,
                    "opening_stock": opening_stock,
                    "closing_stock": closing_stock,
                    "average_stock": average_stock,
                    "sale_count": sale_count,
                    "purchase_count": purchase_count,
                    "turnover_ratio": turnover,
                    "category_id": product.categ_id.complete_name,
                    "company_id": data["company"].name,
                    "warehouse_id": data["warehouse"].name if data["warehouse"] else "",
                }
            else:
                values = {
                    "id": product.id,
                    "opening_stock": opening_stock,
                    "closing_stock": closing_stock,
                    "average_stock": average_stock,
                    "sale_count": sale_count,
                    "purchase_count": purchase_count,
                    "turnover_ratio": turnover,
                    "category_id": product.categ_id.id,
                    "company_id": company_id,
                    "warehouse_id": warehouse_id,
                }

            stock_report.append(values)

        return stock_report

    def _get_batch_stock_at_date(self, product_ids, location_ids, target_date):
        """Batch calculate stock for all products/locations at a specific date using SQL.
        Returns dict: {(product_id, location_id): quantity}
        """
        if not target_date:
            # Use current stock from quants
            self.env.cr.execute(
                """
                SELECT product_id, location_id, SUM(quantity)
                FROM stock_quant
                WHERE product_id IN %s
                AND location_id IN %s
                GROUP BY product_id, location_id
            """,
                (tuple(product_ids), tuple(location_ids)),
            )

            result = {}
            for row in self.env.cr.fetchall():
                result[(row[0], row[1])] = row[2] or 0
            return result

        # Calculate stock at specific date using stock moves
        target_datetime = datetime.combine(target_date, time.max)

        # Query incoming moves (to our locations)
        self.env.cr.execute(
            """
            SELECT 
                sm.product_id,
                sm.location_dest_id as location_id,
                SUM(sm.product_uom_qty) as qty
            FROM stock_move sm
            WHERE sm.product_id IN %s
            AND sm.state = 'done'
            AND sm.date <= %s
            AND sm.location_dest_id IN %s
            GROUP BY sm.product_id, sm.location_dest_id
        """,
            (tuple(product_ids), target_datetime, tuple(location_ids)),
        )

        result = {}
        for row in self.env.cr.fetchall():
            key = (row[0], row[1])
            result[key] = result.get(key, 0) + (row[2] or 0)

        # Query outgoing moves (from our locations)
        self.env.cr.execute(
            """
            SELECT 
                sm.product_id,
                sm.location_id,
                SUM(sm.product_uom_qty) as qty
            FROM stock_move sm
            WHERE sm.product_id IN %s
            AND sm.state = 'done'
            AND sm.date <= %s
            AND sm.location_id IN %s
            GROUP BY sm.product_id, sm.location_id
        """,
            (tuple(product_ids), target_datetime, tuple(location_ids)),
        )

        for row in self.env.cr.fetchall():
            key = (row[0], row[1])
            result[key] = result.get(key, 0) - (row[2] or 0)

        return result

    def _get_batch_sales(self, product_ids, location_ids):
        """Batch get sales quantities for all products/locations in date range.
        Returns dict: {(product_id, location_id): quantity}
        """
        domain = [
            ("product_id", "in", product_ids),
            ("location_id", "in", location_ids),
            ("state", "=", "done"),
            ("location_dest_id.usage", "=", "customer"),
        ]

        if self.start_date:
            start_datetime = datetime.combine(self.start_date, time.min)
            domain.append(("date", ">=", start_datetime))

        if self.end_date:
            end_datetime = datetime.combine(self.end_date, time.max)
            domain.append(("date", "<=", end_datetime))

        # Use read_group for aggregation
        moves_data = self.env["stock.move"].read_group(
            domain, ["product_id", "location_id", "product_uom_qty"], ["product_id", "location_id"], lazy=False
        )

        result = {}
        for data in moves_data:
            product_id = data["product_id"][0] if data["product_id"] else False
            location_id = data["location_id"][0] if data["location_id"] else False
            if product_id and location_id:
                result[(product_id, location_id)] = data["product_uom_qty"] or 0

        return result

    def _get_batch_purchases(self, product_ids, location_ids):
        """Batch get purchase quantities for all products/locations in date range.
        Returns dict: {(product_id, location_id): quantity}
        """
        domain = [
            ("product_id", "in", product_ids),
            ("location_dest_id", "in", location_ids),
            ("state", "=", "done"),
            ("location_id.usage", "=", "supplier"),
        ]

        if self.start_date:
            start_datetime = datetime.combine(self.start_date, time.min)
            domain.append(("date", ">=", start_datetime))

        if self.end_date:
            end_datetime = datetime.combine(self.end_date, time.max)
            domain.append(("date", "<=", end_datetime))

        # Use read_group for aggregation
        moves_data = self.env["stock.move"].read_group(
            domain,
            ["product_id", "location_dest_id", "product_uom_qty"],
            ["product_id", "location_dest_id"],
            lazy=False,
        )

        result = {}
        for data in moves_data:
            product_id = data["product_id"][0] if data["product_id"] else False
            location_id = data["location_dest_id"][0] if data["location_dest_id"] else False
            if product_id and location_id:
                result[(product_id, location_id)] = data["product_uom_qty"] or 0

        return result
