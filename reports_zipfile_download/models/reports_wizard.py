# models/reports_wizard.py
from odoo import models, fields, api
from odoo.exceptions import ValidationError, UserError
from datetime import datetime
import re
import xlsxwriter
import base64
from io import BytesIO
import json
import csv
import threading
import logging
import odoo
import odoo.sql_db

# ============================================================================
# NEW MODEL: Print Label Line
# ============================================================================
class PrintLabelLine(models.TransientModel):
    _name = 'print.label.line'
    _description = 'Print Label Line'
    
    wizard_id = fields.Many2one('reports.wizard', string='Wizard', ondelete='cascade', required=True)
    dc_id = fields.Many2one('stock.picking', string='DC', required=True)
    dc_number = fields.Char(related='dc_id.dc_number', string='DC Number', readonly=True, store=True)
    customer_name = fields.Char(related='dc_id.partner_id.name', string='Customer', readonly=True, store=True)
    
    # FIXED: Changed from fields.Date to fields.Datetime
    scheduled_date = fields.Datetime(related='dc_id.scheduled_date', string='Scheduled Date', readonly=True, store=True)
    
    label_count = fields.Integer(string='Label Count', default=1, required=True)
    
    @api.constrains('label_count')
    def _check_label_count(self):
        for record in self:
            if record.label_count < 1:
                raise ValidationError('Label count must be at least 1')
            if record.label_count > 999:
                raise ValidationError('Label count cannot exceed 999')

# ============================================================================
# MAIN WIZARD MODEL
# ============================================================================
class ReportsWizard(models.TransientModel):
    _name = 'reports.wizard'
    _description = 'Centralized Reports Wizard'



    name = fields.Char(string='Name', compute='_compute_name', store=False)
    
    # Common fields
    report_type = fields.Selection([
        ('overview_report', 'Overview Report'),
        ('enquiry_report', 'Enquiry Report'),
        ('quotation_report', 'Quotation Report'),
        ('workorder_summary', 'Workorder Summary'),
        ('purchase_order_report', 'Purchase Order Report'),
        ('purchase_indent_report', 'Purchase Indent Report'),
        ('consignee_separation', 'Consignee Separation'),
        ('dc_report', 'DC Report'),
        ('delivery_booking_summary_report', 'Delivery Booking Summary Report'),
        ('task_allocation_report', 'Task Allocation Report'),
        ('dc_movement_report', 'DC Movement Report'),
        ('invoice_summary_report', 'Invoice Summary Report'),
        ('installation_summary', 'Installation Summary'),
        ('item_summary', 'Item Summary'),
        ('delivered_undelivered_summary', 'Delivered/Undelivered Summary'),
        ('courier', 'Courier'),
        ('serial_number_tracking', 'Serial Number Product Tracking'),
        ('stock_summary_report', 'Stock Summary Report'),
        ('print_label', 'Print Label'),
        ('quotation_print_report', 'Quotation Print Report'),
        ('dc_print_report', 'DC Print Report'),
        ('po_print_report', 'Purchase Order Print Report'),
        ('invoice_print_report', 'Invoice Print Report'),
        ('installation_print_report', 'Installation Print Report'),
        ('elcot_print_report', 'ELCOT Print Report'),
        ('product_summary', 'Product Summary'),
    ], string='Reports Type', required=True, default='overview_report')

    print_report_type = fields.Selection([
        ('quotation_print_report', 'Quotation Print Report'),
        ('dc_print_report', 'DC Print Report'),
        ('po_print_report', 'Purchase Order Print Report'),
        ('invoice_print_report', 'Invoice Print Report'),
        ('installation_print_report', 'Installation Print Report'),
        ('elcot_print_report', 'ELCOT Print Report'),
    ], string='Reports Type')

    @api.onchange('print_report_type')
    def _onchange_print_report_type(self):
        if self.print_report_type:
            self.report_type = self.print_report_type

    from_date = fields.Date(string='From Date')
    to_date = fields.Date(string='To Date')
    customer_id = fields.Many2one('res.partner', string='Customer Name', domain=[('customer_rank', '>', 0)])
    salesperson_id = fields.Many2one('res.users', string='Salesperson')
    
    # Enquiry Report fields
    enquiry_from_date = fields.Date(string='From Date')
    enquiry_to_date = fields.Date(string='To Date')
    # enquiry_status = fields.Selection([
    #     ('draft', 'Draft'),
    #     ('in_progress', 'In Progress'),
    #     ('won', 'Won'),
    #     ('lost', 'Lost'),
    #     ('cancelled', 'Cancelled'),
    # ], string='Status')
    enquiry_status = fields.Selection([
        ('new', 'New'),
        ('open', 'Open'),
        ('reject', 'Closed'),
        ('contacted', 'Converted To Lead'),
    ], string='Status')
    # enquiry_customer_id = fields.Many2one('res.partner', string='Customer Name')
    # enquiry_customer_name = fields.Char(string='Customer Name')
    enquiry_customer_name = fields.Selection(
        selection='_get_company_selection',
        string='Customer Name'
    )
    # Task Allocation Report fields
    task_from_date = fields.Date(string='From Date')
    task_to_date = fields.Date(string='To Date')
    task_employee_id = fields.Many2one('hr.employee', string='Employee')
    task_project_id = fields.Many2one('project.project', string='Project')
    task_status = fields.Selection([
        ('01_in_progress', 'In Progress'),
        ('02_changes_requested', 'Changes Requested'),
        ('03_approved', 'Approved'),
        ('1_done', 'Done'),
        ('1_canceled', 'Canceled'),
        ('04_waiting_normal', 'Waiting'),
    ], string='Task Status')
    task_stage_id = fields.Many2one(
        'project.task.type', 
        string='Task Stage',
        options="{'no_create': True, 'no_edit': True}"
    )

    @api.model
    def _get_company_selection(self):
        """Get all unique company names for selection dropdown"""
        enquiries = self.env['customer.enq'].search([])
        companies = set()
        
        for enq in enquiries:
            if enq.customer_type == 'new' and enq.company:
                companies.add(enq.company)
            elif enq.customer_type == 'exist' and enq.company_id:
                companies.add(enq.company_id.name)
        
        # Return as list of tuples (value, label)
        return [(name, name) for name in sorted(companies)]
    
    # Quotation Report fields
    # quotation_number = fields.Char(string='Quotation Number')
    quotation_number_ids = fields.Many2many(
        'quotation.management',
        'reports_wizard_quotation_number_rel',
        string='Quotation Numbers'
    )
    enquiry_number = fields.Char(string='Enquiry Number')
    enquiry_ids = fields.Many2many(
        'order.enq',
        string="Enquiry Numbers"
    )
    quotation_from_date = fields.Date(string='From Date')
    quotation_to_date = fields.Date(string='To Date')
    # quotation_status = fields.Selection([
    #     ('draft', 'Draft'),
    #     ('sent', 'Sent'),
    #     ('approved', 'Approved'),
    #     ('rejected', 'Rejected'),
    # ], string='Status')
    quotation_status = fields.Selection([
        ('draft', 'Draft'),
        ('negotiate', 'Negotiation'),
        ('waiting_approval', 'Waiting for Approval'),
        ('rejected', 'Rejected'),
        ('approved', 'Approved'),
        ('quote_sent', 'Quotation Sent'),
        ('resend_for_approval', 'Resend for Approval'),
        ('cancel', 'Cancelled'),
        ('confirm', 'Order Confirmed'),
        ('inactive', 'Inactive'),
    ], string='Status')
    quotation_customer_id = fields.Many2one('res.partner', string='Customer Name')
    
    # Workorder Summary fields
    workorder_number = fields.Char(string='Workorder Number')
    workorder_quotation_number = fields.Char(string='Quotation Number')
    workorder_enquiry_number = fields.Char(string='Enquiry Number')
    workorder_number_ids = fields.Many2many(
        'sale.order',
        'reports_wizard_workorder_number_rel',  # unique relation table name
        string='Workorder Numbers'
    )

    workorder_quotation_number_ids = fields.Many2many(
        'quotation.management',
        'reports_wizard_workorder_quotation_rel',  # unique relation table name
        string='Quotation Numbers'
    )

    workorder_enquiry_number_ids = fields.Many2many(
        'customer.enq',
        'reports_wizard_workorder_enquiry_rel',  # unique relation table name
        string='Enquiry Numbers'
    )
    workorder_from_date = fields.Date(string='From Date')
    workorder_to_date = fields.Date(string='To Date')
    workorder_customer_id = fields.Many2one('res.partner', string='Customer Name')
    
    # Consignee Separation fields
    consignee_workorder_number = fields.Char(string='Workorder Number')
    consignee_enquiry_number = fields.Char(string='Enquiry Number')
    consignee_separation_ids = fields.Many2many(
        'consignee.separation',
        string="Consignee Separation Number"
    )

    consignee_from_date = fields.Date(string='From Date')
    consignee_to_date = fields.Date(string='To Date')
    consignee_customer_id = fields.Many2one('res.partner', string='Customer Name')
    
    # DC Report fields
    dc_number = fields.Char(string='DC Number')
    dc_workorder_number = fields.Char(string='Workorder Number')
    dc_number_ids = fields.Many2many(
        'stock.picking',
        'wizard_dc_number_rel',
        'wizard_id',
        'picking_id',
        string="DC Number"
    )

    dc_workorder_number_ids = fields.Many2many(
        'sale.order',
        string="Workorder Numbers"
    )
    dc_from_date = fields.Date(string='From Date')
    dc_to_date = fields.Date(string='To Date')
    dc_customer_id = fields.Many2one('res.partner', string='Customer Name')
    
    # Item Summary fields
    item_summary_from_date = fields.Date(string='From Date')
    item_summary_to_date = fields.Date(string='To Date')
    item_name = fields.Char(string='Item Name')

    # Product Summary Report filter fields
    product_summary_category_id = fields.Many2one('product.category', string='Product Category')
    product_summary_product_type = fields.Selection([
        ('consu', 'Consumable'),
        ('service', 'Service'),
        ('product', 'Storable Product'),
    ], string='Product Type')
    product_summary_name = fields.Many2one('product.template', string='Product Name')

    # Installation Summary Report fields
    installation_from_date = fields.Date(string='From Date')
    installation_to_date = fields.Date(string='To Date')
    # installation_project_id = fields.Many2one('project.project', string='Project')
    installation_project_id = fields.Many2one(
        'project.project', 
        string='Project',
        domain=[('name', 'ilike', 'INS')]
    )
    installation_project_ids = fields.Many2many(
        'project.project',
        'wizard_installation_project_rel',
        'wizard_id',
        'project_id',
        string="Project",
        domain=[('name', 'ilike', 'INS')]
    )
    installation_customer_id_filter = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    
    # Serial Number Tracking filter fields
    serial_product_id = fields.Many2one('product.product', string='Product',
        domain=[('tracking', '!=', 'none')])

    serial_product_ids = fields.Many2many(
        'product.product',
        'wizard_serial_product_rel',
        'wizard_id',
        'product_id',
        string="Product",
        domain=[('tracking', '!=', 'none')]
    )
    serial_from_date = fields.Date(string='From Date')
    serial_to_date = fields.Date(string='To Date')

    # DC Movement Report filter fields
    dc_movement_from_date = fields.Date(string='From Date')
    dc_movement_to_date = fields.Date(string='To Date')
    dc_movement_dc_number = fields.Many2one(
        'stock.picking', 
        string='DC Number',
        domain=[('picking_type_code', '=', 'outgoing')]
    )
    dc_movement_dc_number_ids = fields.Many2many(
        'stock.picking',
        'wizard_dc_movement_rel',
        'wizard_id',
        'picking_id',
        string="DC Number",
        domain=[('picking_type_code', '=', 'outgoing')]
    )
    dc_movement_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    dc_movement_warehouse_id = fields.Many2one('stock.warehouse', string='Warehouse')
    dc_movement_status = fields.Selection([
        ('draft', 'Draft'),
        ('waiting', 'Waiting Another Operation'),
        ('confirmed', 'Waiting'),
        ('assigned', 'Inspection'),
        ('outward', 'Outward'),
        ('inward', 'Inward'),
        ('done', 'Done'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Courier Report filter fields
    courier_from_date = fields.Date(string='From Date')
    courier_to_date = fields.Date(string='To Date')
    # courier_partner_filter_id = fields.Many2one('courier.partner', string='Courier Name')
    courier_partner_filter_id = fields.Many2one('res.partner', string='Courier Name')
    courier_order_type = fields.Selection([
        ('elcot', 'ELCOT Order'),
        ('others', 'Direct Order'),
    ], string='Order Type')

    # Delivery Booking Summary Report filter fields
    delivery_booking_from_date = fields.Date(string='From Date')
    delivery_booking_to_date = fields.Date(string='To Date')
    delivery_booking_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    delivery_booking_dispatch_type = fields.Selection([
        ('direct', 'Direct'),
        ('courier', 'Courier'),
    ], string='Dispatch Type')
    delivery_booking_status = fields.Selection([
        ('draft', 'Draft'),
        ('confirmed', 'Confirmed'),
        ('dispatched', 'Dispatched'),
        ('delivered', 'Delivered'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Delivered/Undelivered Summary filter fields
    du_from_date = fields.Date(string='From Date')
    du_to_date = fields.Date(string='To Date')
    du_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    du_workorder_ids = fields.Many2many(
        'sale.order',
        'wizard_du_workorder_rel',
        'wizard_id',
        'sale_id',
        string='Workorder Numbers',
    )

    # DC Print Report filter fields
    dc_print_from_date = fields.Date(string='From Date')
    dc_print_to_date = fields.Date(string='To Date')
    dc_print_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    dc_print_warehouse_id = fields.Many2one('stock.warehouse', string='Warehouse')
    dc_print_dc_number = fields.Many2one(
        'stock.picking',
        string='DC Number',
        domain=[('picking_type_code', '=', 'outgoing')],
    )
    dc_print_dc_number_ids = fields.Many2many(
        'stock.picking',
        'wizard_dc_print_rel',
        'wizard_id',
        'picking_id',
        string="DC Number",
        domain=[('picking_type_code', '=', 'outgoing')]
    )
    dc_print_state = fields.Selection([
        ('draft', 'Draft'),
        ('waiting', 'Waiting Another Operation'),
        ('confirmed', 'Waiting'),
        ('assigned', 'Inspection'),
        ('outward', 'Outward'),
        ('done', 'Done'),
        ('cancel', 'Cancelled'),
    ], string='Status')
    print_download_type = fields.Selection([
        ('pdf', 'PDF (Merged)'),
        ('zip', 'ZIP (Separate Files)'),
    ], string='Download Type', default='pdf')

    # Purchase Order Report filter fields
    # purchase_order_number = fields.Char(string='PO Number')
    purchase_order_id = fields.Many2one(
        'purchase.order',
        string='PO Number',
        domain=[],
    )
    purchase_order_ids = fields.Many2many(
        'purchase.order',
        'reports_wizard_purchase_order_rel',
        'wizard_id',
        'purchase_id',
        string="PO Numbers"
    )
    purchase_order_from_date = fields.Date(string='From Date')
    purchase_order_to_date = fields.Date(string='To Date')
    purchase_order_vendor_id = fields.Many2one('res.partner', string='Vendor', domain=[('supplier_rank', '>', 0)])
    purchase_order_salesperson_id = fields.Many2one('res.users', string='Salesperson')
    purchase_order_status = fields.Selection([
        ('draft', 'Draft'),
        ('sent', 'Approval'),
        ('to approve', 'To Approve'),
        ('purchase', 'Purchase Order'),
        ('done', 'Locked'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Purchase Indent Report filter fields
    purchase_indent_id = fields.Many2one(
        'purchase.indent',
        string='Indent No',
        domain=[],
    )
    purchase_indent_ids = fields.Many2many(
        'purchase.indent',
        'reports_wizard_purchase_indent_rel',
        'wizard_id',
        'indent_id',
        string="Indent Numbers"
    )
    purchase_indent_from_date = fields.Date(string='From Date')
    purchase_indent_to_date = fields.Date(string='To Date')
    purchase_indent_vendor_id = fields.Many2one('res.partner', string='Vendor', domain=[('supplier_rank', '>', 0)])
    purchase_indent_order_type = fields.Selection([
        ('workorder', 'Against Work order'),
        ('direct', 'Direct'),
    ], string='Order Type')
    purchase_indent_status = fields.Selection([
        ('draft', 'Draft'),
        ('approved', 'Approved'),
        ('vendor_mapping', 'Vendor Mapping Done'),
        # ('done', 'Done'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Invoice Summary Report filter fields
    invoice_summary_id = fields.Many2one(
        'account.move',
        string='Invoice Number',
        domain=[('move_type', '=', 'out_invoice')],
    )
    invoice_summary_ids = fields.Many2many(
        'account.move',
        'wizard_invoice_summary_rel',
        'wizard_id',
        'move_id',
        string="Invoice Number",
        domain=[('move_type', '=', 'out_invoice')]
    )
    invoice_summary_from_date = fields.Date(string='From Date')
    invoice_summary_to_date = fields.Date(string='To Date')
    invoice_summary_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    invoice_summary_invoice_type = fields.Selection([
        ('customer', 'Customer'),
        ('consignee', 'Consignee'),
    ], string='Invoice Type')
    invoice_summary_status = fields.Selection([
        ('draft', 'Draft'),
        ('posted', 'Posted'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # PO Print Report filter fields
    po_print_id = fields.Many2one(
        'purchase.order',
        string='PO Number',
        domain=[],
    )
    po_print_ids = fields.Many2many(
        'purchase.order',
        'wizard_po_print_rel',
        'wizard_id',
        'po_id',
        string="PO Number"
    )
    po_print_from_date = fields.Date(string='From Date')
    po_print_to_date = fields.Date(string='To Date')
    po_print_vendor_id = fields.Many2one('res.partner', string='Vendor', domain=[('supplier_rank', '>', 0)])
    po_print_status = fields.Selection([
        ('draft', 'Draft'),
        ('sent', 'Approval'),
        ('to approve', 'To Approve'),
        ('purchase', 'Purchase Order'),
        ('done', 'Locked'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Invoice Print Report filter fields
    invoice_print_id = fields.Many2one(
        'account.move',
        string='Invoice Number',
        domain=[('move_type', '=', 'out_invoice')],
    )
    invoice_print_ids = fields.Many2many(
        'account.move',
        'wizard_invoice_print_rel',
        'wizard_id',
        'move_id',
        string="Invoice Number",
        domain=[('move_type','=','out_invoice')]
    )
    invoice_print_from_date = fields.Date(string='From Date')
    invoice_print_to_date = fields.Date(string='To Date')
    invoice_print_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    invoice_print_status = fields.Selection([
        ('draft', 'Draft'),
        ('posted', 'Posted'),
        ('cancel', 'Cancelled'),
    ], string='Status')

    # Quotation Print Report filter fields
    quotation_print_id = fields.Many2one(
        'quotation.management',
        string='Quotation Number',
        domain=[],
    )
    quotation_print_ids = fields.Many2many(
        'quotation.management',
        'wizard_quotation_print_rel',
        'wizard_id',
        'quotation_id',
        string="Quotation Number"
    )
    quotation_print_from_date = fields.Date(string='From Date')
    quotation_print_to_date = fields.Date(string='To Date')
    quotation_print_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    quotation_print_status = fields.Selection([
        ('approved', 'Approved'),
        ('quote_sent', 'Quotation Sent'),
        ('confirm', 'Order Confirmed'),
    ], string='Status')

    # Installation Print Report filter fields
    installation_print_from_date = fields.Date(string='From Date')
    installation_print_to_date = fields.Date(string='To Date')
    installation_print_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    installation_print_project_id = fields.Many2one('project.project', string='Project', domain=[('name', 'ilike', 'INS')])
    installation_print_project_ids = fields.Many2many(
        'project.project',
        'wizard_installation_print_project_rel',
        'wizard_id',
        'project_id',
        string="Project",
        domain=[('name','ilike','INS')]
    )
    installation_print_task_id = fields.Many2one('project.task', string='Task')
    installation_print_task_ids = fields.Many2many(
        'project.task',
        'wizard_installation_print_task_rel',
        'wizard_id',
        'task_id',
        string="Task"
    )
    installation_print_stage_id = fields.Many2one(
        'project.task.type',
        string='Stage',
        options="{'no_create': True, 'no_edit': True}"
    )
    installation_print_state = fields.Selection([
        ('01_in_progress', 'In Progress'),
        ('02_changes_requested', 'Changes Requested'),
        ('03_approved', 'Approved'),
        ('1_done', 'Done'),
        ('1_canceled', 'Canceled'),
    ], string='State')

    # ELCOT Print Report filter fields
    elcot_print_from_date = fields.Date(string='From Date')
    elcot_print_to_date = fields.Date(string='To Date')
    elcot_print_customer_id = fields.Many2one('res.partner', string='Customer', domain=[('customer_rank', '>', 0)])
    elcot_print_project_ids = fields.Many2many(
        'project.project',
        'wizard_elcot_print_project_rel',
        'wizard_id',
        'project_id',
        string="Project",
        domain=[('name', 'ilike', 'INS')]
    )
    elcot_print_task_ids = fields.Many2many(
        'project.task',
        'wizard_elcot_print_task_rel',
        'wizard_id',
        'task_id',
        string="Task"
    )
    elcot_print_state = fields.Selection([
        ('01_in_progress', 'In Progress'),
        ('02_changes_requested', 'Changes Requested'),
        ('03_approved', 'Approved'),
        ('1_done', 'Done'),
        ('1_canceled', 'Canceled'),
    ], string='State')

    # Print Label fields - UPDATED
    # print_label_dc_ids = fields.Many2many(
    #     'stock.picking', 
    #     string='Select DC Numbers',
    #     domain=[('picking_type_code', '=', 'outgoing')]
    # )

    print_label_dc_ids = fields.Many2many(
        'stock.picking',
        'wizard_print_label_dc_rel',
        'wizard_id',
        'picking_id',
        string="Print Label DC"
    )
    print_label_warehouse_id = fields.Many2one(
        'stock.warehouse',
        string='Warehouse',
    )
    print_label_line_ids = fields.One2many(
        'print.label.line', 
        'wizard_id', 
        string='DC Label Lines'
    )
    
    bulk_import_file = fields.Binary(string='Import File (CSV/XLSX)')
    bulk_import_filename = fields.Char(string='Filename')

    # Output fields
    excel_file = fields.Binary(string='Excel File', readonly=True, attachment=False)
    filename = fields.Char(string='Filename')
    pdf_file = fields.Binary(string='PDF File', readonly=True, attachment=False)   # ADD THIS
    pdf_filename = fields.Char(string='PDF Filename') 
    preview_count = fields.Integer(string='Preview Count', compute='_compute_preview_count')
    preview_ids = fields.One2many('report.preview', 'wizard_id', string='Preview Data')
    
    @api.depends('report_type')
    def _compute_name(self):
        """Generate display name based on report type"""
        for record in self:
            report_names = {
                'overview_report': 'Overview Report',
                'enquiry_report': 'Enquiry Report',
                'quotation_report': 'Quotation Report',
                'workorder_summary': 'Workorder Summary',
                'consignee_separation': 'Consignee Separation',
                'dc_report': 'DC Report',
                'task_allocation_report': 'Task Allocation Report',
                'dc_movement_report': 'DC Movement Report',
                'installation_summary': 'Installation Summary',
                'item_summary': 'Item Summary',
                'courier': 'Courier Report',
                'serial_number_tracking': 'Serial Number Tracking',
                'print_label': 'Print Label Generator',
                'product_summary': 'Product Summary',

            }
            record.name = report_names.get(record.report_type, 'Reports Generator')
            
    @api.depends('preview_ids')
    def _compute_preview_count(self):
        for record in self:
            record.preview_count = len(record.preview_ids)
    
    # ========================================================================
    # NEW: BULK IMPORT METHODS
    # ========================================================================
    
    def download_bulk_import_template(self):
        """Generate and download Excel template for bulk DC import"""
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('DC Label Template')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        instruction_format = workbook.add_format({
            'italic': True,
            'font_color': '#666666',
            'text_wrap': True
        })
        
        example_format = workbook.add_format({
            'bg_color': '#E7E6E6',
            'border': 1
        })
        
        
        # Set column widths
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 15)
        
        # Write headers (row 7)
        worksheet.write('A1', 'DC Number', header_format)
        worksheet.write('B1', 'Label Count', header_format)
        
        # Write example data (rows 8-10)
        examples = [
            ('DC00001/25-26', 5),
            ('DC00003/25-26', 10)
        ]
        
        for idx, (dc_num, count) in enumerate(examples, start=2):
            worksheet.write(f'A{idx}', dc_num, example_format)
            worksheet.write(f'B{idx}', count, example_format)
        
        workbook.close()
        output.seek(0)
        
        # Save and return download action
        template_data = base64.b64encode(output.read())
        filename = f'DC_Label_Import_Template_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': template_data,
            'res_model': 'reports.wizard',
            'res_id': self.id,
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'new',
        }
    
    def import_bulk_dc_data(self):
        """Import DC numbers and counts from CSV/Excel file"""
        if not self.bulk_import_file:
            raise UserError('Please upload a file to import')
        
        # Decode file
        file_data = base64.b64decode(self.bulk_import_file)
        
        # Determine file type
        is_excel = self.bulk_import_filename and (
            self.bulk_import_filename.endswith('.xlsx') or 
            self.bulk_import_filename.endswith('.xls')
        )
        
        imported_data = []
        errors = []
        
        try:
            if is_excel:
                # Handle Excel files
                import openpyxl
                from io import BytesIO
                
                workbook = openpyxl.load_workbook(BytesIO(file_data))
                sheet = workbook.active
                
                # Find header row (look for "DC Number" in column A)
                header_row = None
                for row_idx, row in enumerate(sheet.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
                    if row[0] and 'DC Number' in str(row[0]):
                        header_row = row_idx
                        break
                
                if not header_row:
                    raise UserError('Could not find header row with "DC Number" in the Excel file')
                
                # Read data rows
                for row_idx, row in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
                    dc_number = row[0]
                    label_count = row[1]
                    
                    # Skip empty rows
                    if not dc_number:
                        continue
                    
                    # Validate data
                    if not label_count:
                        errors.append(f'Row {row_idx}: Missing label count for DC {dc_number}')
                        continue
                    
                    try:
                        label_count = int(label_count)
                        if label_count < 1 or label_count > 999:
                            errors.append(f'Row {row_idx}: Label count must be between 1 and 999 for DC {dc_number}')
                            continue
                    except (ValueError, TypeError):
                        errors.append(f'Row {row_idx}: Invalid label count "{label_count}" for DC {dc_number}')
                        continue
                    
                    imported_data.append({
                        'dc_number': str(dc_number).strip(),
                        'label_count': label_count,
                        'row': row_idx
                    })
                    
            else:
                # Handle CSV files
                file_content = file_data.decode('utf-8-sig')
                csv_reader = csv.DictReader(file_content.splitlines())
                
                for row_idx, row in enumerate(csv_reader, start=2):
                    dc_number = row.get('DC Number', '').strip()
                    label_count_str = row.get('Label Count', '').strip()
                    
                    if not dc_number:
                        continue
                    
                    if not label_count_str:
                        errors.append(f'Row {row_idx}: Missing label count for DC {dc_number}')
                        continue
                    
                    try:
                        label_count = int(label_count_str)
                        if label_count < 1 or label_count > 999:
                            errors.append(f'Row {row_idx}: Label count must be between 1 and 999 for DC {dc_number}')
                            continue
                    except (ValueError, TypeError):
                        errors.append(f'Row {row_idx}: Invalid label count "{label_count_str}" for DC {dc_number}')
                        continue
                    
                    imported_data.append({
                        'dc_number': dc_number,
                        'label_count': label_count,
                        'row': row_idx
                    })
        
        except Exception as e:
            raise UserError(f'Error reading file: {str(e)}')
        
        if not imported_data and not errors:
            raise UserError('No data found in the import file. Please check the file format.')
        
        # Find matching DC records
        dc_commands = []
        line_commands = []
        
        for data in imported_data:
            dc_number = data['dc_number']
            label_count = data['label_count']
            row = data['row']
            
            # Search for DC by dc_number or name field
            dc = self.env['stock.picking'].search([
                ('picking_type_code', '=', 'outgoing'),
                '|',
                ('dc_number', '=', dc_number),
                ('name', '=', dc_number)
            ], limit=1)
            
            if not dc:
                errors.append(f'Row {row}: DC Number "{dc_number}" not found')
                continue
            
            # Add to many2many and create line
            dc_commands.append((4, dc.id))
            line_commands.append((0, 0, {
                'dc_id': dc.id,
                'label_count': label_count,
            }))
        
        # Update wizard with imported data
        if dc_commands:
            # Clear existing data first
            self.print_label_line_ids = [(5, 0, 0)]
            self.print_label_dc_ids = [(5, 0, 0)]
            
            # Add new data
            self.print_label_dc_ids = dc_commands
            self.print_label_line_ids = line_commands
        
         # Show result message
        message = f'Successfully imported {len(line_commands)} DC records'
        if errors:
            message += f'\n\nErrors ({len(errors)}):\n' + '\n'.join(errors[:10])
            if len(errors) > 10:
                message += f'\n... and {len(errors) - 10} more errors'
        
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Bulk Import Complete',
                'message': message,
                'type': 'warning' if errors else 'success',
                'sticky': False,  # Changed from True to False for auto-close
                'next': {'type': 'ir.actions.act_window_close'},
            }
        }
    
    # ========================================================================
    # PRINT LABEL METHODS - UPDATED
    # ========================================================================
    
    # @api.onchange('print_label_dc_ids')
    # def _onchange_print_label_dc_ids(self):
    #     """Create/update print_label_line_ids when DC selection changes"""
    #     if not self.print_label_dc_ids:
    #         self.print_label_line_ids = [(5, 0, 0)]  # Clear all lines
    #         return
        
    #     # Get existing line DC IDs (handle both NewId and real IDs)
    #     existing_lines = {}
    #     for line in self.print_label_line_ids:
    #         dc_id = line.dc_id.id if line.dc_id else False
    #         if dc_id:
    #             existing_lines[dc_id] = line
        
    #     # Get selected DC IDs (handle both NewId and real IDs)
    #     selected_dc_ids = set()
    #     for dc in self.print_label_dc_ids:
    #         dc_id = dc._origin.id if dc._origin else dc.id
    #         if dc_id:
    #             selected_dc_ids.add(dc_id)
        
    #     commands = []
        
    #     # Remove lines for unselected DCs
    #     for dc_id, line in existing_lines.items():
    #         if dc_id not in selected_dc_ids:
    #             if line.id:
    #                 commands.append((2, line.id, 0))  # Delete existing line
        
    #     # Add lines for newly selected DCs
    #     for dc in self.print_label_dc_ids:
    #         dc_id = dc._origin.id if dc._origin else dc.id
    #         if dc_id and dc_id not in existing_lines:
    #             commands.append((0, 0, {
    #                 'dc_id': dc_id,
    #                 'label_count': 1,
    #             }))
        
    #     if commands:
    #         self.print_label_line_ids = commands

    @api.onchange('print_label_dc_ids')
    def _onchange_print_label_dc_ids(self):
        """Create/update print_label_line_ids when DC selection changes"""
        if not self.print_label_dc_ids:
            self.print_label_line_ids = [(5, 0, 0)]  # Clear all lines
            return

        # Get existing line DC IDs safely
        existing_lines = {}
        for line in self.print_label_line_ids:
            # Use _origin to get the real stored id
            dc = line.dc_id
            real_id = dc._origin.id if hasattr(dc, '_origin') and dc._origin and dc._origin.id else dc.id
            if real_id and not isinstance(real_id, models.NewId):  # ← skip fake ids
                existing_lines[real_id] = line

        # Get selected DC IDs safely
        selected_dc_ids = set()
        for dc in self.print_label_dc_ids:
            real_id = dc._origin.id if hasattr(dc, '_origin') and dc._origin and dc._origin.id else dc.id
            if real_id and not isinstance(real_id, models.NewId):  # ← skip fake ids
                selected_dc_ids.add(real_id)

        commands = []

        # Remove lines for unselected DCs
        for dc_id, line in existing_lines.items():
            if dc_id not in selected_dc_ids:
                if line.id and not isinstance(line.id, models.NewId):
                    commands.append((2, line.id, 0))  # Delete existing saved line
                else:
                    commands.append((3, line.id, 0))  # Unlink unsaved line

        # Add lines for newly selected DCs
        for dc in self.print_label_dc_ids:
            real_id = dc._origin.id if hasattr(dc, '_origin') and dc._origin and dc._origin.id else dc.id
            if real_id and not isinstance(real_id, models.NewId) and real_id not in existing_lines:
                commands.append((0, 0, {
                    'dc_id': real_id,
                    'label_count': 1,
                }))

        if commands:
            self.print_label_line_ids = commands
    
    def generate_print_label_report(self):
        """Generate Print Label PDF"""
        if not self.print_label_line_ids:
            raise ValidationError('Please select at least one DC Number and set label counts')
        
        label_data = []
        all_dc_ids = []
        
        for line in self.print_label_line_ids:
            dc_id = line.dc_id.id
            count = line.label_count
            
            # Validate count
            if count < 1:
                raise ValidationError(f'Label count for DC {line.dc_number} must be at least 1')
            
            # Generate labels for this DC
            for i in range(1, count + 1):
                label_data.append({
                    'dc_id': dc_id,
                    'current_count': i,
                    'total_count': count,
                })
            
            all_dc_ids.append(dc_id)
        
        # Generate PDF
        # repeated_pickings = self.env['stock.picking'].browse(all_dc_ids)
        # return self.env.ref('hrms_dashboard.action_report_print_label_multi').with_context(
        #     label_data=label_data
        # ).report_action(repeated_pickings)

        # Generate PDF
        repeated_pickings = self.env['stock.picking'].browse(all_dc_ids)
        return self.env.ref('hrms_dashboard.action_report_print_label_multi').with_context(
            label_data=label_data,
            selected_warehouse_id=self.print_label_warehouse_id.id if self.print_label_warehouse_id else False,
        ).report_action(repeated_pickings)
    
    # ========================================================================
    # MAIN GENERATE & PREVIEW METHODS (Keep these as is)
    # ========================================================================
    
    def generate_report(self):
        """Main method to route to specific report generator"""
        if self.print_report_type:
            self.report_type = self.print_report_type
        self._validate_report_params()

        # Clear previous files before generating new report
        self.write({'pdf_file': False, 'pdf_filename': False, 
                    'excel_file': False, 'filename': False})
        
        if self.report_type == 'overview_report':
            return self.generate_overview_report()
        elif self.report_type == 'enquiry_report':
            return self.generate_enquiry_report()
        elif self.report_type == 'quotation_report':
            return self.generate_quotation_report()
        elif self.report_type == 'workorder_summary':
            return self.generate_workorder_report()
        elif self.report_type == 'consignee_separation':
            return self.generate_consignee_report()
        elif self.report_type == 'dc_report':
            return self.generate_dc_report()
        elif self.report_type == 'task_allocation_report':
            return self.generate_task_allocation_report()
        elif self.report_type == 'item_summary':
            return self.generate_item_summary_report()
        elif self.report_type == 'installation_summary':
            return self.generate_installation_summary_report()
        elif self.report_type == 'serial_number_tracking':
            return self.generate_serial_number_report()
        elif self.report_type == 'stock_summary_report':
            return self.generate_stock_summary_report()
        elif self.report_type == 'dc_movement_report':
            return self.generate_dc_movement_report()
        elif self.report_type == 'courier':
            return self.generate_courier_report()
        elif self.report_type == 'print_label':
            return self.generate_print_label_report()
        elif self.report_type == 'delivery_booking_summary_report':
            return self.generate_delivery_booking_summary_report()
        elif self.report_type == 'delivered_undelivered_summary':
            return self.generate_delivered_undelivered_summary_report()
        elif self.report_type == 'dc_print_report':
            return self.generate_dc_print_report()
        elif self.report_type == 'purchase_order_report':
            return self.generate_purchase_order_report()
        elif self.report_type == 'po_print_report':
            return self.generate_po_print_report()
        elif self.report_type == 'purchase_indent_report':
            return self.generate_purchase_indent_report()
        elif self.report_type == 'invoice_summary_report':
            return self.generate_invoice_summary_report()
        elif self.report_type == 'invoice_print_report':
            return self.generate_invoice_print_report()
        elif self.report_type == 'quotation_print_report':
            return self.generate_quotation_print_report()
        elif self.report_type == 'installation_print_report':
            return self.generate_installation_print_report()
        elif self.report_type == 'elcot_print_report':
            return self.generate_elcot_print_report()
        elif self.report_type == 'product_summary':
            return self.generate_product_summary_report()
        else:
            raise ValidationError(f'Report type {self.report_type} is not yet implemented')
    
    def generate_preview(self):
        """Generate preview data for the selected report"""
        self._validate_report_params()
        
        # Clear existing preview
        self.preview_ids.unlink()
        
        if self.report_type == 'overview_report':
            preview_data = self._get_overview_preview_data()
        elif self.report_type == 'enquiry_report':
            preview_data = self._get_enquiry_preview_data()
        elif self.report_type == 'quotation_report':
            preview_data = self._get_quotation_preview_data()
        elif self.report_type == 'workorder_summary':
            preview_data = self._get_workorder_preview_data()
        elif self.report_type == 'consignee_separation':
            preview_data = self._get_consignee_preview_data()
        elif self.report_type == 'dc_report':
            preview_data = self._get_dc_preview_data()
        elif self.report_type == 'task_allocation_report':
            preview_data = self._get_task_allocation_preview_data()
        elif self.report_type == 'installation_summary':
            preview_data = self._get_installation_summary_preview_data()
        elif self.report_type == 'item_summary':
            preview_data = self._get_item_summary_preview_data()
        elif self.report_type == 'serial_number_tracking':
            preview_data = self._get_serial_number_preview_data()
        elif self.report_type == 'stock_summary_report':
            preview_data = self._get_stock_summary_preview_data()
        elif self.report_type == 'dc_movement_report':
            preview_data = self._get_dc_movement_preview_data()
        elif self.report_type == 'purchase_order_report':
            preview_data = self._get_purchase_order_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'po_number': data.get('po_number', ''),
                    'po_vendor': data.get('po_vendor', ''),
                    'po_salesperson': data.get('po_salesperson', ''),
                    'po_created_date': data.get('po_created_date') or False,
                    'po_expected_arrival': data.get('po_expected_arrival') or False,
                    'po_status': data.get('po_status', ''),
                    'po_purchase_indent': data.get('po_purchase_indent', ''),
                    'po_work_order_no': data.get('po_work_order_no', ''),
                    'po_customer_po_number': data.get('po_customer_po_number', ''),
                    'po_po_issue_date': data.get('po_po_issue_date') or False,
                    'po_grn_numbers': data.get('po_grn_numbers', ''),
                    'po_grn_date': data.get('po_grn_date', ''),
                    'po_grn_status': data.get('po_grn_status', ''),
                    'po_purchase_person': data.get('po_purchase_person', ''),
                    'po_product': data.get('po_product', ''),
                    'po_description': data.get('po_description', ''),
                    'po_qty_ordered': data.get('po_qty_ordered', 0),
                    'po_qty_received': data.get('po_qty_received', 0),
                    'po_qty_billed': data.get('po_qty_billed', 0),
                    'po_uom': data.get('po_uom', ''),
                    'po_unit_price': data.get('po_unit_price', 0.0),
                    'po_subtotal': data.get('po_subtotal', 0.0),
                    'po_total': data.get('po_total', 0.0),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }
        elif self.report_type == 'purchase_indent_report':
            preview_data = self._get_purchase_indent_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'pi_indent_no': data.get('pi_indent_no', ''),
                    'pi_indent_date': data.get('pi_indent_date') or False,
                    'pi_order_type': data.get('pi_order_type', ''),
                    'pi_sale_order': data.get('pi_sale_order', ''),
                    'pi_approved_date': data.get('pi_approved_date') or False,
                    'pi_vendor': data.get('pi_vendor', ''),
                    'pi_status': data.get('pi_status', ''),
                    'pi_product': data.get('pi_product', ''),
                    'pi_part_number': data.get('pi_part_number', ''),
                    'pi_qty': data.get('pi_qty', 0.0),
                    'pi_price': data.get('pi_price', 0.0),
                    'pi_uom': data.get('pi_uom', ''),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }

        elif self.report_type == 'invoice_summary_report':
            preview_data = self._get_invoice_summary_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'inv_number': data.get('inv_number', ''),
                    'inv_customer': data.get('inv_customer', ''),
                    'inv_date': data.get('inv_date') or False,
                    'inv_due_date': data.get('inv_due_date') or False,
                    'inv_dc_number': data.get('inv_dc_number', ''),
                    'inv_work_order': data.get('inv_work_order', ''),
                    'inv_po_number': data.get('inv_po_number', ''),
                    'inv_po_issue_date': data.get('inv_po_issue_date') or False,
                    'inv_invoice_type': data.get('inv_invoice_type', ''),
                    'inv_status': data.get('inv_status', ''),
                    'inv_item': data.get('inv_item', ''),
                    'inv_item_name': data.get('inv_item_name', ''),
                    'inv_description': data.get('inv_description', ''),
                    'inv_qty': data.get('inv_qty', 0),
                    'inv_taxes': data.get('inv_taxes', ''),
                    'inv_subtotal': data.get('inv_subtotal', 0.0),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }
        elif self.report_type == 'courier':
            preview_data = self._get_courier_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'courier_dc_number': data.get('courier_dc_number', ''),
                    'courier_sale_number': data.get('courier_sale_number', ''),
                    'courier_order_type_val': data.get('courier_order_type_val', ''),
                    'courier_customer': data.get('courier_customer', ''),
                    'courier_customer_po_number': data.get('courier_customer_po_number', ''),
                    'courier_po_issue_date': data.get('courier_po_issue_date') or False,
                    'courier_booking_date_field': data.get('courier_booking_date_field') or False,
                    'courier_booking_date_val': data.get('courier_booking_date_val') or False,
                    'courier_dc_date': data.get('courier_dc_date') or False,
                    'courier_partner_name': data.get('courier_partner_name', ''),
                    'courier_booking_number': data.get('courier_booking_number', ''),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }
        elif self.report_type == 'delivery_booking_summary_report':
            preview_data = self._get_delivery_booking_summary_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'db_booking_number': data.get('db_booking_number', ''),
                    'db_dc_number': data.get('db_dc_number', ''),
                    'db_sale_number': data.get('db_sale_number', ''),
                    'db_customer_po_number': data.get('db_customer_po_number', ''),
                    'db_po_issue_date': data.get('db_po_issue_date') or False,
                    'db_dispatch_location': data.get('db_dispatch_location', ''),
                    'db_customer': data.get('db_customer', ''),
                    'db_billing_address': data.get('db_billing_address', ''),
                    'db_shipping_address': data.get('db_shipping_address', ''),
                    'db_contact_person': data.get('db_contact_person', ''),
                    'db_contact_no': data.get('db_contact_no', ''),
                    'db_booking_date': data.get('db_booking_date') or False,
                    'db_dispatch_type': data.get('db_dispatch_type', ''),
                    'db_surface': data.get('db_surface', ''),
                    'db_driver_name': data.get('db_driver_name', ''),
                    'db_vehicle_number': data.get('db_vehicle_number', ''),
                    'db_status': data.get('db_status', ''),
                    'db_district': data.get('db_district', ''),
                    'db_item_details': data.get('db_item_details', ''),
                    'db_item_description': data.get('db_item_description', ''),
                    'db_item_qty': data.get('db_item_qty', ''),
                    'db_unit_price': data.get('db_unit_price', 0.0),
                    'db_dc_value': data.get('db_dc_value', 0.0),
                    'db_invoice_no': data.get('db_invoice_no', ''),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }
        elif self.report_type == 'delivered_undelivered_summary':
            preview_data = self._get_delivered_undelivered_preview_data()
            for data in preview_data[:50]:
                self.env['report.preview'].create({
                    'wizard_id': self.id,
                    'du_workorder_no': data.get('du_workorder_no', ''),
                    'du_customer_po_number': data.get('du_customer_po_number', ''),
                    'du_po_issue_date': data.get('du_po_issue_date') or False,
                    'du_customer': data.get('du_customer', ''),
                    'du_order_type': data.get('du_order_type', ''),
                    'du_item_name': data.get('du_item_name', ''),
                    'du_item_description': data.get('du_item_description', ''),
                    'du_workorder_qty': data.get('du_workorder_qty', 0.0),
                    'du_unit_price': data.get('du_unit_price', 0.0),
                    'du_total_amount': data.get('du_total_amount', 0.0),
                    'du_delivered_qty': data.get('du_delivered_qty', 0.0),
                    'du_delivered_amount': data.get('du_delivered_amount', 0.0),
                    'du_undelivered_qty': data.get('du_undelivered_qty', 0.0),
                    'du_undelivered_amount': data.get('du_undelivered_amount', 0.0),
                    'du_inst_completed_qty': data.get('du_inst_completed_qty', 0.0),
                    'du_inst_completed_value': data.get('du_inst_completed_value', 0.0),
                    'du_inst_incomplete_qty': data.get('du_inst_incomplete_qty', 0.0),
                    'du_inst_incomplete_value': data.get('du_inst_incomplete_value', 0.0),
                })
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': 'Preview Generated',
                    'message': f'Showing first 50 records out of {len(preview_data)} total records',
                    'type': 'success',
                    'sticky': False,
                }
            }
        else:
            raise ValidationError(f'Preview for {self.report_type} is not yet implemented')
        
        # Create preview records (limit to 50)
        for data in preview_data[:50]:
            data['wizard_id'] = self.id
            self.env['report.preview'].create(data)
        
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Preview Generated',
                'message': f'Showing first 50 records out of {len(preview_data)} total records',
                'type': 'success',
                'sticky': False,
            }
        }
    
    def _validate_report_params(self):
        """Validate date ranges and required fields"""
        if self.report_type in ['overview_report', 'enquiry_report', 'quotation_report', 
                                'workorder_summary', 'consignee_separation', 'dc_report']:
            from_date_field = f"{self.report_type.replace('_report', '').replace('overview', 'from')}_from_date"
            to_date_field = f"{self.report_type.replace('_report', '').replace('overview', 'to')}_to_date"
            
            if self.report_type == 'overview_report':
                from_date_field = 'from_date'
                to_date_field = 'to_date'
            
            from_date = getattr(self, from_date_field, None)
            to_date = getattr(self, to_date_field, None)
            
            if from_date and to_date and from_date > to_date:
                raise ValidationError('From Date cannot be greater than To Date')
    
    # ========================================================================
    # Keep all your existing report generation methods below
    # (generate_overview_report, generate_enquiry_report, etc.)
    # ========================================================================
    
    # ... [REST OF YOUR CODE - ALL OTHER METHODS REMAIN THE SAME] ...

    
    
    # def _get_quotation_preview_data(self):
    #     """Get preview data - one row per product with item data repeated + total summary row"""
    #     domain = []
        
    #     # Build domain filters
    #     if self.quotation_from_date:
    #         from_datetime = fields.Datetime.to_datetime(self.quotation_from_date)
    #         domain.append(('quotation_date', '>=', from_datetime))
        
    #     if self.quotation_to_date:
    #         to_datetime = fields.Datetime.to_datetime(self.quotation_to_date).replace(hour=23, minute=59, second=59)
    #         domain.append(('quotation_date', '<=', to_datetime))
        
    #     if self.quotation_number:
    #         domain.append(('name', 'ilike', self.quotation_number))
        
    #     if self.enquiry_number:
    #         domain.append(('enq_id.name', 'ilike', self.enquiry_number))
        
    #     if self.quotation_customer_id:
    #         domain.append(('customer_id', '=', self.quotation_customer_id.id))
        
    #     if self.quotation_status != 'all':
    #         domain.append(('state', '=', self.quotation_status))
        
    #     quotations = self.env['quotation.management'].search(domain, order='quotation_date asc')
        
    #     preview_data = []
        
    #     for quotation in quotations:
    #         # Common quotation data
    #         common_data = {
    #             'quotation_number_field': quotation.name,
    #             'quotation_status': dict(quotation._fields['state'].selection).get(quotation.state, quotation.state),
    #             'quotation_customer': quotation.customer_id.name if quotation.customer_id else '',
    #             'contact_person_name': quotation.contact_person.name if quotation.contact_person else '',
    #             'contact_email': quotation.contact_email or '',
    #             'quotation_date_field': quotation.quotation_date,
    #         }
            
    #         # Calculate totals for this quotation
    #         total_item_amount = 0
    #         total_item_tax = 0
    #         total_product_amount = 0
    #         total_product_tax = 0
    #         total_margin_amount = 0
            
    #         # Process each item
    #         for item_line in quotation.item_line_ids:
    #             # Item warranty info
    #             item_warranty = ''
    #             if item_line.is_warranty == 'yes':
    #                 item_warranty = f"Zigma Warranty: {item_line.period or 'N/A'}"
                
    #             # Item data (will be repeated for each product)
    #             item_data = {
    #                 'item_name': item_line.item_name or '',
    #                 'item_description': item_line.description or '',
    #                 'item_quantity': item_line.quotation_qty,
    #                 'item_unit_price': item_line.price_unit,
    #                 'item_tax_amount': item_line.price_tax,
    #                 'item_total_amount': item_line.price_total,
    #                 'item_hsn_code': item_line.hsn_code or '',
    #                 'item_warranty': item_warranty,
    #             }
                
    #             # Add to totals
    #             total_item_amount += item_line.price_total
    #             total_item_tax += item_line.price_tax
                
    #             # Get linked products
    #             linked_products = self.env['quotation.product.line'].search([
    #                 ('quotation_id', '=', quotation.id),
    #                 ('item_details_id', '=', item_line.item_details_id.id)
    #             ])
                
    #             # Create one row per product
    #             for product in linked_products:
    #                 # Product warranty info
    #                 product_warranty = ''
    #                 if product.is_warranty == 'yes':
    #                     product_warranty = f"Product Warranty: {product.periods or 'N/A'}"
                    
    #                 # Product data (unique for each row)
    #                 product_data = {
    #                     'product_name': product.product_id.name if product.product_id else '',
    #                     'product_quantity': product.product_uom_qty,
    #                     'product_unit_price': product.price_unit,
    #                     'product_tax_amount': product.price_tax,
    #                     'product_total_amount': product.price_total,
    #                     'product_hsn_code': product.product_id.l10n_in_hsn_code or '',
    #                     'product_warranty': product_warranty,
    #                     'product_margin_percentage': product.margin_percentage,
    #                     'product_margin_amount': product.margin,
    #                 }
                    
    #                 # Add to totals
    #                 total_product_amount += product.price_total
    #                 total_product_tax += product.price_tax
    #                 total_margin_amount += product.margin
                    
    #                 # Combine all data for this row
    #                 row_data = {}
    #                 row_data.update(common_data)
    #                 row_data.update(item_data)  # Item data repeated
    #                 row_data.update(product_data)  # Product data unique
                    
    #                 preview_data.append(row_data)
            
    #         # NEW: Add TOTAL SUMMARY ROW for this quotation
    #         total_row = {}
    #         total_row.update(common_data)
    #         total_row.update({
    #             # Item columns for total row
    #             'item_name': '** TOTAL **',
    #             'item_description': 'Summary for quotation',
    #             'item_quantity': 0,
    #             'item_unit_price': 0,
    #             'item_tax_amount': total_item_tax,
    #             'item_total_amount': total_item_amount,
    #             'item_hsn_code': '',
    #             'item_warranty': '',
                
    #             # Product columns for total row
    #             'product_name': '** GRAND TOTAL **',
    #             'product_quantity': 0,
    #             'product_unit_price': 0,
    #             'product_tax_amount': total_product_tax,
    #             'product_total_amount': total_product_amount,  # This will be â‚¹71,715.00
    #             'product_hsn_code': '',
    #             'product_warranty': '',
    #             'product_margin_percentage': 0,
    #             'product_margin_amount': total_margin_amount,
    #         })
            
    #         preview_data.append(total_row)
        
    #     return preview_data







    
    # def generate_quotation_report(self):
    #     """Generate Quotation Report - one row per product with separate item/product columns + totals"""
    #     import datetime
    #     import base64
    #     from io import BytesIO
    #     import xlsxwriter

    #     # Build domain based on filters
    #     domain = []
        
    #     if self.quotation_from_date:
    #         from_datetime = fields.Datetime.to_datetime(self.quotation_from_date)
    #         domain.append(('quotation_date', '>=', from_datetime))
        
    #     if self.quotation_to_date:
    #         to_datetime = fields.Datetime.to_datetime(self.quotation_to_date).replace(hour=23, minute=59, second=59)
    #         domain.append(('quotation_date', '<=', to_datetime))
        
    #     if self.quotation_number:
    #         domain.append(('name', 'ilike', self.quotation_number))
        
    #     if self.enquiry_number:
    #         domain.append(('enq_id.name', 'ilike', self.enquiry_number))
        
    #     if self.quotation_customer_id:
    #         domain.append(('customer_id', '=', self.quotation_customer_id.id))
        
    #     if self.quotation_status != 'all':
    #         domain.append(('state', '=', self.quotation_status))

    #     quotations = self.env['quotation.management'].search(domain, order='quotation_date asc')

    #     # Create Excel file
    #     output = BytesIO()
    #     workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    #     worksheet = workbook.add_worksheet('Quotation Report')

    #     # Define formats
    #     header_format = workbook.add_format({
    #         'bold': True, 
    #         'bg_color': '#4472C4', 
    #         'font_color': 'white', 
    #         'border': 1,
    #         'text_wrap': True, 
    #         'valign': 'vcenter', 
    #         'align': 'center'
    #     })
        
    #     normal_format = workbook.add_format({
    #         'border': 1, 
    #         'text_wrap': True, 
    #         'valign': 'vcenter'
    #     })
        
    #     number_format = workbook.add_format({
    #         'num_format': '#,##0.00', 
    #         'border': 1, 
    #         'valign': 'vcenter'
    #     })
        
    #     date_format = workbook.add_format({
    #         'num_format': 'dd/mm/yyyy', 
    #         'border': 1, 
    #         'valign': 'vcenter'
    #     })
        
    #     total_format = workbook.add_format({
    #         'bold': True, 
    #         'border': 1, 
    #         'bg_color': '#FFD700', 
    #         'valign': 'vcenter'
    #     })
        
    #     total_number_format = workbook.add_format({
    #         'bold': True, 
    #         'border': 1, 
    #         'bg_color': '#FFD700', 
    #         'num_format': '#,##0.00', 
    #         'valign': 'vcenter'
    #     })

    #     # Column headers with separate item and product columns
    #     columns = [
    #         'Quotation #', 'Status', 'Customer', 'Contact Person', 'Contact Email', 'Quotation Date',
    #         'Item Name', 'Item Description', 'Item Quantity', 'Item Unit Price', 'Item Tax', 'Item Total', 'Item HSN', 'Item Warranty',
    #         'Product Name', 'Product Quantity', 'Product Unit Price', 'Product Tax', 'Product Total', 'Product HSN', 'Product Warranty',
    #         'Product Margin %', 'Product Margin Amount'
    #     ]

    #     # Set column widths
    #     widths = [20, 12, 20, 20, 25, 18, 25, 30, 12, 15, 15, 15, 15, 20, 30, 12, 15, 15, 15, 20, 20, 15, 20]
    #     for idx, width in enumerate(widths):
    #         worksheet.set_column(idx, idx, width)

    #     # Write headers
    #     for col_num, col_title in enumerate(columns):
    #         worksheet.write(0, col_num, col_title, header_format)

    #     row = 1

    #     for quotation in quotations:
    #         status = dict(quotation._fields['state'].selection).get(quotation.state, quotation.state)

    #         # Initialize totals for this quotation
    #         total_item_amount = 0.0
    #         total_item_tax = 0.0
    #         total_product_amount = 0.0
    #         total_product_tax = 0.0
    #         total_margin_amount = 0.0

    #         # Process each item line
    #         for item_line in quotation.item_line_ids:
    #             total_item_amount += item_line.price_total
    #             total_item_tax += item_line.price_tax
                
    #             # Get linked products for this item
    #             linked_products = self.env['quotation.product.line'].search([
    #                 ('quotation_id', '=', quotation.id),
    #                 ('item_details_id', '=', item_line.item_details_id.id)
    #             ])

    #             # Create one row per product (with item data repeated)
    #             for product in linked_products:
    #                 total_product_amount += product.price_total
    #                 total_product_tax += product.price_tax
    #                 total_margin_amount += product.margin

    #                 # Item warranty info
    #                 item_warranty = ''
    #                 if item_line.is_warranty == 'yes':
    #                     item_warranty = f"Zigma Warranty: {item_line.period or 'N/A'}"

    #                 # Product warranty info
    #                 product_warranty = ''
    #                 if product.is_warranty == 'yes':
    #                     product_warranty = f"Product Warranty: {product.periods or 'N/A'}"

    #                 # Data row with separate item and product columns
    #                 data = [
    #                     # Common columns
    #                     quotation.name,
    #                     status,
    #                     quotation.customer_id.name if quotation.customer_id else '',
    #                     quotation.contact_person.name if quotation.contact_person else '',
    #                     quotation.contact_email if quotation.contact_email else '',
    #                     quotation.quotation_date,
                        
    #                     # ITEM COLUMNS (repeated for all products of this item)
    #                     item_line.item_name or '',
    #                     item_line.description or '',
    #                     item_line.quotation_qty,
    #                     item_line.price_unit,
    #                     item_line.price_tax,
    #                     item_line.price_total,
    #                     item_line.hsn_code or '',
    #                     item_warranty,
                        
    #                     # PRODUCT COLUMNS (unique for each product)
    #                     product.product_id.name if product.product_id else '',
    #                     product.product_uom_qty,
    #                     product.price_unit,
    #                     product.price_tax,
    #                     product.price_total,
    #                     product.product_id.l10n_in_hsn_code if product.product_id else '',
    #                     product_warranty,
    #                     product.margin_percentage,
    #                     product.margin
    #                 ]

    #                 # Write data row with appropriate formatting
    #                 for col_idx, val in enumerate(data):
    #                     fmt = normal_format
    #                     if col_idx == 5:  # Quotation Date
    #                         fmt = date_format
    #                     elif col_idx in [9, 10, 11, 15, 16, 17, 21, 22]:  # Price/tax/total/margin columns
    #                         fmt = number_format
    #                     worksheet.write(row, col_idx, val, fmt)

    #                 row += 1

    #         # NEW: Write TOTAL SUMMARY ROW for this quotation
    #         total_values = [
    #             # Common columns
    #             quotation.name,
    #             status,
    #             quotation.customer_id.name if quotation.customer_id else '',
    #             quotation.contact_person.name if quotation.contact_person else '',
    #             quotation.contact_email if quotation.contact_email else '',
    #             quotation.quotation_date,
                
    #             # ITEM COLUMNS (totals)
    #             '** TOTAL **', 
    #             'Summary for quotation', 
    #             0, 
    #             0, 
    #             total_item_tax, 
    #             total_item_amount, 
    #             '', 
    #             '',
                
    #             # PRODUCT COLUMNS (totals)
    #             '** GRAND TOTAL **', 
    #             0, 
    #             0, 
    #             total_product_tax, 
    #             total_product_amount,  # This will be â‚¹71,715.00
    #             '', 
    #             '', 
    #             0, 
    #             total_margin_amount
    #         ]

    #         # Write total row with gold background formatting
    #         for col_idx, val in enumerate(total_values):
    #             fmt = total_format
    #             if col_idx == 5:  # Quotation Date
    #                 fmt = date_format
    #             elif col_idx in [10, 11, 17, 18, 22]:  # Amount columns
    #                 fmt = total_number_format
    #             worksheet.write(row, col_idx, val, fmt)

    #         row += 2  # Add blank row between quotations

    #     workbook.close()
    #     output.seek(0)

    #     # Save the file with dynamic name
    #     filter_parts = []
    #     if self.quotation_from_date and self.quotation_to_date:
    #         filter_parts.append(f'{self.quotation_from_date}_{self.quotation_to_date}')
    #     if self.quotation_customer_id:
    #         filter_parts.append(f'customer_{self.quotation_customer_id.name.replace(" ", "_")}')
        
    #     filename = f'quotation_report_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
    #     self.write({
    #         'excel_file': base64.b64encode(output.read()),
    #         'filename': filename
    #     })
        
    #     # Generate preview after creating Excel
    #     self.generate_preview()
        
    #     return {
    #         'type': 'ir.actions.act_window',
    #         'res_model': 'reports.wizard',
    #         'view_mode': 'form',
    #         'res_id': self.id,
    #         'target': 'current',
    #     }

    def _get_quotation_preview_data(self):
        """Get preview data - one row per product with item data repeated"""
        domain = []

        if self.quotation_from_date:
            domain.append(('quotation_date', '>=', self.quotation_from_date))

        if self.quotation_to_date:
            domain.append(('quotation_date', '<=', self.quotation_to_date))

        # if self.quotation_number:
        #     domain.append(('name', 'ilike', self.quotation_number))
        if self.quotation_number_ids:
            domain.append(('id', 'in', self.quotation_number_ids.ids))

        # if self.enquiry_number:
        #     domain.append(('enq_id.name', 'ilike', self.enquiry_number))

        if self.enquiry_ids:
            domain.append(('lead', 'in', self.enquiry_ids.ids))

        if self.quotation_customer_id:
            domain.append(('customer_id', '=', self.quotation_customer_id.id))

        if self.quotation_status and self.quotation_status != 'all':
            domain.append(('state', '=', self.quotation_status))

        quotations = self.env['quotation.management'].search(domain, order='quotation_date asc')

        preview_data = []

        for quotation in quotations:
            status_dict = dict(quotation._fields['state'].selection)
            status = status_dict.get(quotation.state, quotation.state)

            common_data = {
                'quotation_number_field': quotation.name or '',
                'quotation_status': status,
                'quotation_customer': quotation.customer_id.name if quotation.customer_id else '',
                'contact_person_name': quotation.contact_person.name if quotation.contact_person else '',
                'contact_email': quotation.contact_email or '',
                'quotation_date_field': quotation.quotation_date if quotation.quotation_date else '',
            }

            # Handle quotations with no item lines
            if not quotation.item_line_ids:
                row_data = {}
                row_data.update(common_data)
                row_data.update({
                    'item_name': '(No items)',
                    'item_description': '',
                    'item_quantity': 0,
                    'item_unit_price': 0.0,
                    'item_tax_amount': 0.0,
                    'item_total_amount': 0.0,
                    'item_hsn_code': '',
                    # 'item_warranty': '',
                    'product_name': '(No products)',
                    'product_quantity': 0,
                    'product_unit_price': 0.0,
                    'product_tax_amount': 0.0,
                    'product_total_amount': 0.0,
                    'product_hsn_code': '',
                    'product_warranty': '',
                })
                preview_data.append(row_data)
                continue

            for item_line in quotation.item_line_ids:

                item_warranty = ''
                if item_line.is_warranty == 'yes':
                    item_warranty = f"Zigma Warranty: {item_line.period or 'N/A'}"

                item_data = {
                    'item_name': item_line.item_name or '',
                    'item_description': item_line.description or '',
                    'item_quantity': item_line.quotation_qty or 0,
                    'item_unit_price': item_line.price_unit or 0.0,
                    'item_tax_amount': item_line.price_tax or 0.0,
                    'item_total_amount': item_line.price_total or 0.0,
                    'item_hsn_code': item_line.hsn_code or '',
                    'item_warranty': item_warranty,
                }

                linked_products = quotation.product_line_ids.filtered(
                    lambda p: p.item_details_id == item_line.item_details_id
                )

                if linked_products:
                    for product in linked_products:
                        product_warranty = product.period or ''

                        product_data = {
                            'product_name': product.product_id.name if product.product_id else '',
                            'product_quantity': product.product_uom_qty or 0,
                            'product_unit_price': product.price_unit or 0.0,
                            'product_tax_amount': product.price_tax or 0.0,
                            'product_total_amount': product.price_total or 0.0,
                            # 'product_hsn_code': product.product_id.l10n_in_hsn_code if product.product_id else '',
                            'product_warranty': product_warranty,
                        }

                        row_data = {}
                        row_data.update(common_data)
                        row_data.update(item_data)
                        row_data.update(product_data)
                        preview_data.append(row_data)
                else:
                    empty_product = {
                        'product_name': '(No products)',
                        'product_quantity': 0,
                        'product_unit_price': 0.0,
                        'product_tax_amount': 0.0,
                        'product_total_amount': 0.0,
                        'product_hsn_code': '',
                        'product_warranty': '',
                    }
                    row_data = {}
                    row_data.update(common_data)
                    row_data.update(item_data)
                    row_data.update(empty_product)
                    preview_data.append(row_data)

        return preview_data


    def generate_quotation_report(self):
        """Generate Quotation Report Excel - one row per product"""
        import datetime
        import base64
        from io import BytesIO
        import xlsxwriter

        domain = []

        if self.quotation_from_date:
            domain.append(('quotation_date', '>=', self.quotation_from_date))

        if self.quotation_to_date:
            domain.append(('quotation_date', '<=', self.quotation_to_date))

        # if self.quotation_number:
        #     domain.append(('name', 'ilike', self.quotation_number))

        if self.quotation_number_ids:
            domain.append(('id', 'in', self.quotation_number_ids.ids))

        # if self.enquiry_number:
        #     domain.append(('enq_id.name', 'ilike', self.enquiry_number))

        if self.enquiry_ids:
            domain.append(('lead', 'in', self.enquiry_ids.ids))

        if self.quotation_customer_id:
            domain.append(('customer_id', '=', self.quotation_customer_id.id))

        if self.quotation_status and self.quotation_status != 'all':
            domain.append(('state', '=', self.quotation_status))

        quotations = self.env['quotation.management'].search(domain, order='quotation_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Quotation Report')

        # ── Formats ──────────────────────────────────────────────
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
        })
        normal_fmt = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter'
        })
        number_fmt = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'
        })
        date_fmt = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter'
        })

        # ── Columns ───────────────────────────────────────────────
        columns = [
            'Quotation', 'Status', 'Customer', 'Contact Person', 'Contact Email', 'Quotation Date',
            'Item Name', 'Item Description', 'Item Quantity',
            'Item Unit Price', 'Item Tax', 'Item Total', 'Item HSN',
            'Item Margin %', 'Item Margin Amount',
            'Product Name', 'Product Quantity', 'Product Unit Price',
            'Product Tax', 'Product Total','Product Warranty', 'Part Number',
        ]
        widths = [
            20, 15, 22, 22, 28, 15,
            28, 32, 10,
            15, 15, 15, 15,
            14, 18,
            30, 12, 15,
            15, 15, 22, 18,
        ]

        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
        for col, title in enumerate(columns):
            worksheet.write(0, col, title, header_fmt)

        # Number-formatted column indexes
        number_cols = {9, 10, 11, 13, 14, 17, 18, 19}

        # ── FIX: write_row defined HERE before any loop ───────────
        def write_row(ws, r, values):
            for c, val in enumerate(values):
                if c == 5:  # Quotation Date column
                    fmt = date_fmt
                elif c in number_cols:
                    fmt = number_fmt
                else:
                    fmt = normal_fmt
                ws.write(r, c, val, fmt)

        # row = 1

        # for quotation in quotations:
        #     status_dict = dict(quotation._fields['state'].selection)
        #     status = status_dict.get(quotation.state, quotation.state)

        #     # ── Handle quotations with no item lines ──────────────
        #     if not quotation.item_line_ids:
        #         data = [
        #             quotation.name or '',
        #             status,
        #             quotation.customer_id.name if quotation.customer_id else '',
        #             quotation.contact_person.name if quotation.contact_person else '',
        #             quotation.contact_email or '',
        #             quotation.quotation_date,
        #             '(No items)', '', 0, 0.0, 0.0, 0.0, '', '',
        #             0.0,
        #             '(No products)', 0.0, 0, 0.0, 0.0, 0.0, '', 
        #         ]
        #         write_row(worksheet, row, data)
        #         row += 1
        #         continue

        #     # ── Process item lines ────────────────────────────────
        #     for item_line in quotation.item_line_ids:

        #         item_warranty = ''
        #         if item_line.is_warranty == 'yes':
        #             item_warranty = f"Zigma Warranty: {item_line.period or 'N/A'}"

        #         linked_products = quotation.product_line_ids.filtered(
        #             lambda p: p.item_details_id == item_line.item_details_id
        #         )

        #         if linked_products:
        #             for product in linked_products:
        #                 product_warranty = product.period or ''

        #                 data = [
        #                     # Common columns
        #                     quotation.name or '',
        #                     status,
        #                     quotation.customer_id.name if quotation.customer_id else '',
        #                     quotation.contact_person.name if quotation.contact_person else '',
        #                     quotation.contact_email or '',
        #                     quotation.quotation_date,           # col 5 → date_fmt
        #                     # Item columns
        #                     item_line.item_name or '',
        #                     item_line.description or '',
        #                     item_line.quotation_qty or 0,
        #                     item_line.price_unit or 0.0,        # col 9
        #                     item_line.price_tax or 0.0,         # col 10
        #                     item_line.price_total or 0.0,       # col 11
        #                     item_line.hsn_code or '',
        #                     # item_warranty,
        #                     item_line.margin_percentage or 0.0, # col 14
        #                     item_line.margin or 0.0,            # col 15
        #                     # Product columns
        #                     product.product_id.name if product.product_id else '',
        #                     product.product_uom_qty or 0,
        #                     product.price_unit or 0.0,          # col 18
        #                     product.price_tax or 0.0,           # col 19
        #                     product.price_total or 0.0,         # col 20
        #                     # product.product_id.l10n_in_hsn_code if product.product_id else '',
        #                     product_warranty,
        #                     product.part_number or '',
        #                 ]
        #                 write_row(worksheet, row, data)
        #                 row += 1

        #         else:
        #             # Item has no linked products - still show item row
        #             data = [
        #                 quotation.name or '',
        #                 status,
        #                 quotation.customer_id.name if quotation.customer_id else '',
        #                 quotation.contact_person.name if quotation.contact_person else '',
        #                 quotation.contact_email or '',
        #                 quotation.quotation_date,
        #                 item_line.item_name or '',
        #                 item_line.description or '',
        #                 item_line.quotation_qty or 0,
        #                 item_line.price_unit or 0.0,
        #                 item_line.price_tax or 0.0,
        #                 item_line.price_total or 0.0,
        #                 item_line.hsn_code or '',
        #                 # item_warranty,
        #                 item_line.margin_percentage or 0.0,
        #                 item_line.margin or 0.0,
        #                 '(No products linked)', 0, 0.0, 0.0, 0.0, '', '',
        #             ]
        #             write_row(worksheet, row, data)
        #             row += 1

        row = 1

        for quotation in quotations:

            start_row = row

            status_dict = dict(quotation._fields['state'].selection)
            status = status_dict.get(quotation.state, quotation.state)

            # Handle quotations with no item lines
            if not quotation.item_line_ids:
                data = [
                    quotation.name or '',
                    status,
                    quotation.customer_id.name if quotation.customer_id else '',
                    quotation.contact_person.name if quotation.contact_person else '',
                    quotation.contact_email or '',
                    quotation.quotation_date,
                    '(No items)',
                    '',
                    0,
                    0.0,
                    0.0,
                    0.0,
                    '',
                    0.0,
                    0.0,
                    '(No products)',
                    0,
                    0.0,
                    0.0,
                    0.0,
                    '',
                    '',
                ]

                write_row(worksheet, row, data)
                row += 1
                continue

            # Count how many rows this quotation will occupy
            total_rows = 0
            for item_line in quotation.item_line_ids:
                linked_products = quotation.product_line_ids.filtered(
                    lambda p: p.item_details_id == item_line.item_details_id
                )
                total_rows += len(linked_products) if linked_products else 1

            if total_rows == 0:
                total_rows = 1

            end_row = start_row + total_rows - 1

            # Write product rows (leave quotation columns empty)
            for item_line in quotation.item_line_ids:

                linked_products = quotation.product_line_ids.filtered(
                    lambda p: p.item_details_id == item_line.item_details_id
                )

                if linked_products:
                    for product in linked_products:
                        data = [
                            '', status,
                            quotation.customer_id.name if quotation.customer_id else '',
                            quotation.contact_person.name if quotation.contact_person else '',
                            quotation.contact_email or '',
                            quotation.quotation_date,

                            item_line.item_name or '',
                            item_line.description or '',
                            item_line.quotation_qty or 0,
                            item_line.price_unit or 0.0,
                            item_line.price_tax or 0.0,
                            item_line.price_total or 0.0,
                            item_line.hsn_code or '',
                            item_line.margin_percentage or 0.0,
                            item_line.margin or 0.0,

                            product.product_id.name if product.product_id else '',
                            product.product_uom_qty or 0,
                            product.price_unit or 0.0,
                            product.price_tax or 0.0,
                            product.price_total or 0.0,
                            product.period or '',
                            product.part_number or '',
                        ]

                        write_row(worksheet, row, data)
                        row += 1
                else:
                    data = [
                        '', status,
                            quotation.customer_id.name if quotation.customer_id else '',
                            quotation.contact_person.name if quotation.contact_person else '',
                            quotation.contact_email or '',
                            quotation.quotation_date,

                        item_line.item_name or '',
                        item_line.description or '',
                        item_line.quotation_qty or 0,
                        item_line.price_unit or 0.0,
                        item_line.price_tax or 0.0,
                        item_line.price_total or 0.0,
                        item_line.hsn_code or '',
                        item_line.margin_percentage or 0.0,
                        item_line.margin or 0.0,

                        '(No products linked)', 0, 0.0, 0.0, 0.0, '', '',
                    ]

                    write_row(worksheet, row, data)
                    row += 1

            # Merge quotation columns vertically
            # worksheet.merge_range(start_row, 0, end_row, 0, quotation.name or '', normal_fmt)
            if start_row == end_row:
                worksheet.write(start_row, 0, quotation.name or '', normal_fmt)
            else:
                worksheet.merge_range(start_row, 0, end_row, 0, quotation.name or '', normal_fmt)
            # worksheet.merge_range(start_row, 1, end_row, 1, status, normal_fmt)
            # worksheet.merge_range(start_row, 2, end_row, 2,
            #                     quotation.customer_id.name if quotation.customer_id else '',
            #                     normal_fmt)
            # worksheet.merge_range(start_row, 3, end_row, 3,
            #                     quotation.contact_person.name if quotation.contact_person else '',
            #                     normal_fmt)
            # worksheet.merge_range(start_row, 4, end_row, 4,
            #                     quotation.contact_email or '',
            #                     normal_fmt)
            # worksheet.merge_range(start_row, 5, end_row, 5,
            #                     quotation.quotation_date,
            #                     date_fmt)

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.quotation_from_date and self.quotation_to_date:
            filter_parts.append(f'{self.quotation_from_date}_{self.quotation_to_date}')
        if self.quotation_customer_id:
            filter_parts.append(f'customer_{self.quotation_customer_id.name.replace(" ", "_")}')

        filename = (
            f'quotation_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }
    
    def _get_workorder_preview_data(self):
        """Get preview data for workorder summary report"""
        domain = []
        
        # Build domain filters
        # if self.workorder_number:
        #     domain.append(('name', 'ilike', self.workorder_number))
        
        # if self.workorder_quotation_number:
        #     domain.append(('quotation_id.name', 'ilike', self.workorder_quotation_number))

        if self.workorder_number_ids:
            domain.append(('id', 'in', self.workorder_number_ids.ids))

        if self.workorder_quotation_number_ids:
            domain.append(('quotation_id', 'in', self.workorder_quotation_number_ids.ids))

        if self.workorder_enquiry_number_ids:
            domain.append(('quotation_id.enq_id', 'in', self.workorder_enquiry_number_ids.ids))
        
        if self.workorder_from_date:
            domain.append(('date_order', '>=', self.workorder_from_date))
        
        if self.workorder_to_date:
            domain.append(('date_order', '<=', self.workorder_to_date))
        
        if self.workorder_customer_id:
            domain.append(('partner_id', '=', self.workorder_customer_id.id))
        
        # if self.workorder_enquiry_number:
        #     domain.append(('quotation_id.enq_id.name', 'ilike', self.workorder_enquiry_number))
        
        sale_orders = self.env['sale.order'].search(domain, order='date_order asc')
        
        preview_data = []
        
        for sale_order in sale_orders:
            # Get GST treatment info from l10n_in_gst_treatment field
            gst_treatment = ''
            if hasattr(sale_order, 'l10n_in_gst_treatment') and sale_order.l10n_in_gst_treatment:
                # Convert selection value to display text
                gst_treatment_dict = dict(sale_order._fields['l10n_in_gst_treatment'].selection)
                gst_treatment = gst_treatment_dict.get(sale_order.l10n_in_gst_treatment, sale_order.l10n_in_gst_treatment)
            elif sale_order.fiscal_position_id:
                gst_treatment = sale_order.fiscal_position_id.name
            else:
                gst_treatment = 'Regular'  # Default value

            
            # Get tax based on info 
            tax_based_on = 'Customer Location' if sale_order.partner_shipping_id else 'Billing Address'
            
            # Process each item line (NOT product lines)
            for item_line in sale_order.item_line_ids:
                preview_data.append({
                    'workorder_number': sale_order.name,
                    'workorder_customer_po_number': sale_order.client_order_ref or '',
                    'workorder_po_issue_date': sale_order.po_issue_date if sale_order.po_issue_date else False,
                    'workorder_customer': sale_order.partner_id.name if sale_order.partner_id else '',
                    'workorder_contact_person': sale_order.contact_person.name if sale_order.contact_person else '',
                    'workorder_source_quotation': sale_order.quotation_id.name if sale_order.quotation_id else '',
                    'workorder_order_type': dict(sale_order._fields['order_type'].selection).get(sale_order.order_type, sale_order.order_type),
                    'workorder_gst_treatment': gst_treatment,
                    'workorder_created_date': sale_order.date_order,
                    'workorder_tax_based_on': tax_based_on,
                    'workorder_item_description': item_line.description or '',
                    'workorder_item_name': item_line.item_name or '',
                    'workorder_individual_price': item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
                    'workorder_hsn_code': item_line.hsn_code or '',
                    'workorder_quotation_quantity': item_line.quotation_qty,
                    'workorder_tax_amount': item_line.price_tax,
                    'workorder_total_amount': item_line.price_total,
                })
        
        return preview_data

    def generate_workorder_report(self):
        """Generate Workorder Summary Report Excel file"""
        domain = []
        
        # Build domain filters (same as preview)
        # if self.workorder_number:
        #     domain.append(('name', 'ilike', self.workorder_number))
        
        # if self.workorder_quotation_number:
        #     domain.append(('quotation_id.name', 'ilike', self.workorder_quotation_number))

        if self.workorder_number_ids:
            domain.append(('id', 'in', self.workorder_number_ids.ids))

        if self.workorder_quotation_number_ids:
            domain.append(('quotation_id', 'in', self.workorder_quotation_number_ids.ids))

        if self.workorder_enquiry_number_ids:
            domain.append(('quotation_id.enq_id', 'in', self.workorder_enquiry_number_ids.ids))
        
        if self.workorder_from_date:
            domain.append(('date_order', '>=', self.workorder_from_date))
        
        if self.workorder_to_date:
            domain.append(('date_order', '<=', self.workorder_to_date))
        
        if self.workorder_customer_id:
            domain.append(('partner_id', '=', self.workorder_customer_id.id))
        
        # if self.workorder_enquiry_number:
        #     domain.append(('quotation_id.enq_id.name', 'ilike', self.workorder_enquiry_number))
        
        sale_orders = self.env['sale.order'].search(domain, order='date_order asc')
        
        # Create Excel file
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Workorder Summary')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'valign': 'vcenter'
        })
        
        # Headers
        headers = [
            'WorkOrder No', 'Customer PO Number', 'PO Issue Date', 'Customer', 'Contact Person', 'Source Quotation', 'Order Type',
            'GST Treatment', 'Created Date', 'Tax Based On', 'Item Description', 'Item Name',
            'Individual Price', 'HSN Code', 'Quotation Quantity', 'Tax Amount', 'Total Amount'
        ]

        # Set column widths
        widths = [20, 22, 15, 25, 20, 20, 15, 20, 15, 20, 40, 25, 15, 15, 15, 15, 15]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
        
        # Write headers
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        # row = 1
        # for sale_order in sale_orders:
        #     # Get GST treatment
        #     gst_treatment = sale_order.fiscal_position_id.name if sale_order.fiscal_position_id else ''
            
        #     # Get tax based on
        #     tax_based_on = 'Customer Location' if sale_order.partner_shipping_id else 'Billing Address'
            
        #     # Process each item line
        #     for item_line in sale_order.item_line_ids:
        #         data = [
        #             sale_order.name,
        #             sale_order.partner_id.name if sale_order.partner_id else '',
        #             sale_order.contact_person.name if sale_order.contact_person else '',
        #             sale_order.quotation_id.name if sale_order.quotation_id else '',
        #             dict(sale_order._fields['order_type'].selection).get(sale_order.order_type, sale_order.order_type),
        #             gst_treatment,
        #             sale_order.date_order,
        #             tax_based_on,
        #             item_line.description or '',
        #             item_line.item_name or '',
        #             item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
        #             item_line.hsn_code or '',
        #             item_line.quotation_qty,
        #             item_line.price_tax,
        #             item_line.price_total,
        #         ]
                
        #         # Write data with appropriate formatting
        #         for col, val in enumerate(data):
        #             fmt = cell_format
        #             if col == 6:  # Created Date
        #                 fmt = date_format
        #             elif col in [10, 13, 14]:  # Price, tax, total columns
        #                 fmt = number_format
        #             worksheet.write(row, col, val, fmt)
                
        #         row += 1

        # Write data
        row = 1
        for sale_order in sale_orders:
            start_row = row

            # Count how many rows this sale order will occupy
            total_rows = len(sale_order.item_line_ids) if sale_order.item_line_ids else 1
            end_row = start_row + total_rows - 1

            gst_treatment = sale_order.fiscal_position_id.name if sale_order.fiscal_position_id else ''

            tax_based_on = 'Customer Location' if sale_order.partner_shipping_id else 'Billing Address'

            customer_po = sale_order.client_order_ref or ''
            po_issue_date = sale_order.po_issue_date if sale_order.po_issue_date else ''

            if sale_order.item_line_ids:
                for item_line in sale_order.item_line_ids:
                    data = [
                        '',  # Leave blank — filled by merge_range below (WorkOrder No)
                        '',  # Leave blank — filled by merge_range below (Customer PO Number)
                        '',  # Leave blank — filled by merge_range below (PO Issue Date)
                        sale_order.partner_id.name if sale_order.partner_id else '',
                        sale_order.contact_person.name if sale_order.contact_person else '',
                        sale_order.quotation_id.name if sale_order.quotation_id else '',
                        dict(sale_order._fields['order_type'].selection).get(sale_order.order_type, sale_order.order_type),
                        gst_treatment,
                        sale_order.date_order,
                        tax_based_on,
                        item_line.description or '',
                        item_line.item_name or '',
                        item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
                        item_line.hsn_code or '',
                        item_line.quotation_qty,
                        item_line.price_tax,
                        item_line.price_total,
                    ]

                    for col, val in enumerate(data):
                        fmt = cell_format
                        if col == 8:     # Created Date
                            fmt = date_format
                        elif col in [12, 15, 16]:  # Individual Price, Tax Amount, Total
                            fmt = number_format
                        worksheet.write(row, col, val, fmt)

                    row += 1
            else:
                # No item lines - write a single blank row
                data = [
                    '', '', '',
                    sale_order.partner_id.name if sale_order.partner_id else '',
                    sale_order.contact_person.name if sale_order.contact_person else '',
                    sale_order.quotation_id.name if sale_order.quotation_id else '',
                    dict(sale_order._fields['order_type'].selection).get(sale_order.order_type, sale_order.order_type),
                    gst_treatment,
                    sale_order.date_order,
                    tax_based_on,
                    '', '', 0, '', 0, 0, 0,
                ]
                for col, val in enumerate(data):
                    fmt = cell_format
                    if col == 8:
                        fmt = date_format
                    elif col in [12, 15, 16]:
                        fmt = number_format
                    worksheet.write(row, col, val, fmt)
                row += 1

            # Merge WorkOrder No (col 0), Customer PO Number (col 1), PO Issue Date (col 2) vertically
            if end_row > start_row:
                worksheet.merge_range(start_row, 0, end_row, 0, sale_order.name or '', cell_format)
                worksheet.merge_range(start_row, 1, end_row, 1, customer_po, cell_format)
                worksheet.merge_range(start_row, 2, end_row, 2, po_issue_date, date_format if po_issue_date else cell_format)
            else:
                worksheet.write(start_row, 0, sale_order.name or '', cell_format)
                worksheet.write(start_row, 1, customer_po, cell_format)
                worksheet.write(start_row, 2, po_issue_date, date_format if po_issue_date else cell_format)
        
        workbook.close()
        output.seek(0)
        
        # Save file
        filter_parts = []
        if self.workorder_from_date and self.workorder_to_date:
            filter_parts.append(f'{self.workorder_from_date}_{self.workorder_to_date}')
        if self.workorder_customer_id:
            filter_parts.append(f'customer_{self.workorder_customer_id.name.replace(" ", "_")}')
        
        filename = f'workorder_summary_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename
        })
        
        self.generate_preview()
        
        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_consignee_preview_data(self):
        """Get preview data for consignee separation report"""
        domain = []
        
        # Build domain filters on sale.order first
        sale_order_domain = []
        
        # if self.consignee_workorder_number:
        #     sale_order_domain.append(('name', 'ilike', self.consignee_workorder_number))

        if self.consignee_separation_ids:
            domain.append(('id', 'in', self.consignee_separation_ids.ids))
        
        if self.consignee_from_date:
            sale_order_domain.append(('date_order', '>=', self.consignee_from_date))
        
        if self.consignee_to_date:
            sale_order_domain.append(('date_order', '<=', self.consignee_to_date))
        
        if self.consignee_customer_id:
            sale_order_domain.append(('partner_id', '=', self.consignee_customer_id.id))
        
        # if self.consignee_enquiry_number:
        #     sale_order_domain.append(('quotation_id.enq_id.name', 'ilike', self.consignee_enquiry_number))
        
        # First get sale orders matching filters
        sale_orders = self.env['sale.order'].search(sale_order_domain, order='date_order asc')
        
        # Then get consignee separations for these sale orders
        # consignee_domain = [('sale_order_id', 'in', sale_orders.ids)]
        consignee_domain = [('sale_order_id', 'in', sale_orders.ids)]

        if self.consignee_separation_ids:
            consignee_domain.append(('id', 'in', self.consignee_separation_ids.ids))
        consignee_separations = self.env['consignee.separation'].search(consignee_domain, order='consignee_date asc')
        
        preview_data = []
        
        for consignee in consignee_separations:
            # Get consignee location address (split address components)
            location_address = ''
            location_code = ''
            if consignee.location_id:
                address_parts = []
                if consignee.location_id.street:
                    address_parts.append(consignee.location_id.street)
                if consignee.location_id.street2:
                    address_parts.append(consignee.location_id.street2)
                if consignee.location_id.city:
                    address_parts.append(consignee.location_id.city)
                if consignee.location_id.state_id:
                    address_parts.append(consignee.location_id.state_id.name)
                if consignee.location_id.zip:
                    address_parts.append(consignee.location_id.zip)
                if consignee.location_id.country_id:
                    address_parts.append(consignee.location_id.country_id.name)
                
                location_address = ', '.join(address_parts)
                location_code = consignee.location_id.name or ''
            
            # Get customer with address
            customer_address = ''
            if consignee.customer_id:
                customer_parts = [consignee.customer_id.name]
                address_components = []
                if consignee.customer_id.street:
                    address_components.append(consignee.customer_id.street)
                if consignee.customer_id.street2:
                    address_components.append(consignee.customer_id.street2)
                if consignee.customer_id.city:
                    address_components.append(consignee.customer_id.city)
                if consignee.customer_id.state_id:
                    address_components.append(consignee.customer_id.state_id.name)
                if consignee.customer_id.zip:
                    address_components.append(consignee.customer_id.zip)
                if consignee.customer_id.country_id:
                    address_components.append(consignee.customer_id.country_id.name)
                
                if address_components:
                    customer_parts.append(', '.join(address_components))
                customer_address = ' - '.join(customer_parts)
            
            # Get Customer PO Number from sale order
            customer_po_number = consignee.sale_order_id.client_order_ref or ''
            
            # Process each item line (NOT product lines) from consignee
            for item_line in consignee.item_line_ids:
                preview_data.append({
                    'consignee_number': consignee.name,
                    'consignee_location_address': location_address,
                    'consignee_location_code': location_code,
                    'consignee_customer_address': customer_address,
                    'consignee_contact_person': consignee.contact_person.name if consignee.contact_person else '',
                    'consignee_workorder_date': consignee.sale_order_id.date_order,
                    'consignee_customer_po_number': customer_po_number,
                    'consignee_po_issue_date': consignee.sale_order_id.po_issue_date if consignee.sale_order_id and consignee.sale_order_id.po_issue_date else False,
                    'consignee_item_description': item_line.description or '',
                    'consignee_item_name': item_line.item_name or '',
                    'consignee_individual_price': item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
                    'consignee_hsn_code': item_line.hsn_code or '',
                    'consignee_quotation_quantity': item_line.total_qty,
                    'consignee_tax_amount': item_line.price_tax,
                    'consignee_total_amount': item_line.price_total,
                })
        
        return preview_data
    
    def generate_consignee_report(self):
        """Generate Consignee Separation Report Excel file"""

        domain = []

        # Build domain filters (same as preview)
        sale_order_domain = []
        
        # if self.consignee_workorder_number:
        #     sale_order_domain.append(('name', 'ilike', self.consignee_workorder_number))

        if self.consignee_separation_ids:
            domain.append(('id', 'in', self.consignee_separation_ids.ids))
        
        if self.consignee_from_date:
            sale_order_domain.append(('date_order', '>=', self.consignee_from_date))
        
        if self.consignee_to_date:
            sale_order_domain.append(('date_order', '<=', self.consignee_to_date))
        
        if self.consignee_customer_id:
            sale_order_domain.append(('partner_id', '=', self.consignee_customer_id.id))
        
        # if self.consignee_enquiry_number:
        #     sale_order_domain.append(('quotation_id.enq_id.name', 'ilike', self.consignee_enquiry_number))
        
        # Get sale orders and their consignee separations
        sale_orders = self.env['sale.order'].search(sale_order_domain, order='date_order asc')
        # consignee_domain = [('sale_order_id', 'in', sale_orders.ids)]
        consignee_domain = [('sale_order_id', 'in', sale_orders.ids)]

        if self.consignee_separation_ids:
            consignee_domain.append(('id', 'in', self.consignee_separation_ids.ids))
        consignee_separations = self.env['consignee.separation'].search(consignee_domain, order='consignee_date asc')
        
        # Create Excel file
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Consignee Separation')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'valign': 'vcenter'
        })
        
        # Headers
        headers = [
            'Consignee No', 'Consignee Address', 'Location Code', 'Customer with Address',
            'Contact Person', 'Workorder Date', 'Customer PO Number', 'PO Issue Date',
            'Item Description', 'Item Name', 'Individual Price', 'HSN Code', 'Quotation Quantity',
            'Tax Amount', 'Total Amount'
        ]

        # Set column widths
        widths = [20, 40, 20, 40, 20, 15, 20, 15, 40, 25, 15, 15, 15, 15, 15]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
        
        # Write headers
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        # row = 1
        # for consignee in consignee_separations:
        #     # Build location address
        #     location_address = ''
        #     location_code = ''
        #     if consignee.location_id:
        #         address_parts = []
        #         if consignee.location_id.street:
        #             address_parts.append(consignee.location_id.street)
        #         if consignee.location_id.street2:
        #             address_parts.append(consignee.location_id.street2)
        #         if consignee.location_id.city:
        #             address_parts.append(consignee.location_id.city)
        #         if consignee.location_id.state_id:
        #             address_parts.append(consignee.location_id.state_id.name)
        #         if consignee.location_id.zip:
        #             address_parts.append(consignee.location_id.zip)
        #         if consignee.location_id.country_id:
        #             address_parts.append(consignee.location_id.country_id.name)
                
        #         location_address = ', '.join(address_parts)
        #         location_code = consignee.location_id.name or ''
            
        #     # Build customer address
        #     customer_address = ''
        #     if consignee.customer_id:
        #         customer_parts = [consignee.customer_id.name]
        #         address_components = []
        #         if consignee.customer_id.street:
        #             address_components.append(consignee.customer_id.street)
        #         if consignee.customer_id.street2:
        #             address_components.append(consignee.customer_id.street2)
        #         if consignee.customer_id.city:
        #             address_components.append(consignee.customer_id.city)
        #         if consignee.customer_id.state_id:
        #             address_components.append(consignee.customer_id.state_id.name)
        #         if consignee.customer_id.zip:
        #             address_components.append(consignee.customer_id.zip)
        #         if consignee.customer_id.country_id:
        #             address_components.append(consignee.customer_id.country_id.name)
                
        #         if address_components:
        #             customer_parts.append(', '.join(address_components))
        #         customer_address = ' - '.join(customer_parts)
            
        #     # Get Customer PO Number
        #     customer_po_number = consignee.sale_order_id.client_order_ref or ''
            
        #     # Process each item line
        #     for item_line in consignee.item_line_ids:
        #         data = [
        #             consignee.name,
        #             location_address,
        #             location_code,
        #             customer_address,
        #             consignee.contact_person.name if consignee.contact_person else '',
        #             consignee.sale_order_id.date_order,
        #             customer_po_number,
        #             item_line.description or '',
        #             item_line.item_name or '',
        #             item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
        #             item_line.hsn_code or '',
        #             item_line.total_qty,
        #             item_line.price_tax,
        #             item_line.price_total,
        #         ]
                
        #         # Write data with appropriate formatting
        #         for col, val in enumerate(data):
        #             fmt = cell_format
        #             if col == 5:  # Workorder Date
        #                 fmt = date_format
        #             elif col in [9, 12, 13]:  # Price, tax, total columns
        #                 fmt = number_format
        #             worksheet.write(row, col, val, fmt)
                
        #         row += 1

        # Write data
        row = 1
        for consignee in consignee_separations:
            start_row = row

            # Build location address
            location_address = ''
            location_code = ''
            if consignee.location_id:
                address_parts = []
                if consignee.location_id.street:
                    address_parts.append(consignee.location_id.street)
                if consignee.location_id.street2:
                    address_parts.append(consignee.location_id.street2)
                if consignee.location_id.city:
                    address_parts.append(consignee.location_id.city)
                if consignee.location_id.state_id:
                    address_parts.append(consignee.location_id.state_id.name)
                if consignee.location_id.zip:
                    address_parts.append(consignee.location_id.zip)
                if consignee.location_id.country_id:
                    address_parts.append(consignee.location_id.country_id.name)
                location_address = ', '.join(address_parts)
                location_code = consignee.location_id.name or ''

            # Build customer address
            customer_address = ''
            if consignee.customer_id:
                customer_parts = [consignee.customer_id.name]
                address_components = []
                if consignee.customer_id.street:
                    address_components.append(consignee.customer_id.street)
                if consignee.customer_id.street2:
                    address_components.append(consignee.customer_id.street2)
                if consignee.customer_id.city:
                    address_components.append(consignee.customer_id.city)
                if consignee.customer_id.state_id:
                    address_components.append(consignee.customer_id.state_id.name)
                if consignee.customer_id.zip:
                    address_components.append(consignee.customer_id.zip)
                if consignee.customer_id.country_id:
                    address_components.append(consignee.customer_id.country_id.name)
                if address_components:
                    customer_parts.append(', '.join(address_components))
                customer_address = ' - '.join(customer_parts)

            customer_po_number = consignee.sale_order_id.client_order_ref or ''
            consignee_po_issue_date = consignee.sale_order_id.po_issue_date if consignee.sale_order_id and consignee.sale_order_id.po_issue_date else ''

            if consignee.item_line_ids:
                for item_line in consignee.item_line_ids:
                    data = [
                        '',  # Column A - filled by merge below
                        location_address,
                        location_code,
                        customer_address,
                        consignee.contact_person.name if consignee.contact_person else '',
                        consignee.sale_order_id.date_order,
                        customer_po_number,
                        consignee_po_issue_date,
                        item_line.description or '',
                        item_line.item_name or '',
                        item_line.price_subtotal / item_line.order_qty if item_line.order_qty else 0,
                        item_line.hsn_code or '',
                        item_line.total_qty,
                        item_line.price_tax,
                        item_line.price_total,
                    ]
                    for col, val in enumerate(data):
                        fmt = cell_format
                        if col in {5, 7}:
                            fmt = date_format
                        elif col in [10, 13, 14]:
                            fmt = number_format
                        worksheet.write(row, col, val, fmt)
                    row += 1
            else:
                data = [
                    '',
                    location_address,
                    location_code,
                    customer_address,
                    consignee.contact_person.name if consignee.contact_person else '',
                    consignee.sale_order_id.date_order,
                    customer_po_number,
                    consignee_po_issue_date,
                    '', '', 0, '', 0, 0, 0,
                ]
                for col, val in enumerate(data):
                    fmt = cell_format
                    if col in {5, 7}:
                        fmt = date_format
                    elif col in [10, 13, 14]:
                        fmt = number_format
                    worksheet.write(row, col, val, fmt)
                row += 1

            end_row = row - 1

            # Merge Consignee No column (column 0) vertically
            if end_row > start_row:
                worksheet.merge_range(start_row, 0, end_row, 0, consignee.name or '', cell_format)
            else:
                worksheet.write(start_row, 0, consignee.name or '', cell_format)
        
        workbook.close()
        output.seek(0)
        
        # Save file
        filter_parts = []
        if self.consignee_from_date and self.consignee_to_date:
            filter_parts.append(f'{self.consignee_from_date}_{self.consignee_to_date}')
        if self.consignee_customer_id:
            filter_parts.append(f'customer_{self.consignee_customer_id.name.replace(" ", "_")}')
        
        filename = f'consignee_separation_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename
        })
        
        self.generate_preview()
        
        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }
    
    def _get_dc_preview_data(self):
        """Get preview data for DC report"""
        domain = []
        
        # Build domain filters on stock.picking for outgoing deliveries only
        domain.append(('picking_type_id.code', '=', 'outgoing'))
        domain.append(('state', '!=', 'cancel'))
        
        # if self.dc_number:
        #     domain.append(('dc_number', 'ilike', self.dc_number))

        if self.dc_number_ids:
            domain.append(('id', 'in', self.dc_number_ids.ids))

        if self.dc_workorder_number_ids:
            domain.append(('sale_id', 'in', self.dc_workorder_number_ids.ids))
        
        if self.dc_from_date:
            domain.append(('scheduled_date', '>=', self.dc_from_date))
        
        if self.dc_to_date:
            domain.append(('scheduled_date', '<=', self.dc_to_date))
        
        if self.dc_customer_id:
            domain.append(('partner_id', '=', self.dc_customer_id.id))
        
        # if self.dc_workorder_number:
        #     # Filter by workorder number through sale_order_id
        #     domain.append(('sale_id.name', 'ilike', self.dc_workorder_number))
        
        stock_pickings = self.env['stock.picking'].search(domain, order='scheduled_date asc')
        
        preview_data = []
        
        for picking in stock_pickings:
            # Get delivery address from partner_id
            delivery_address = ''
            district = ''
            state_name = ''
            if picking.partner_id:
                address_parts = []
                if picking.partner_id.street:
                    address_parts.append(picking.partner_id.street)
                if picking.partner_id.street2:
                    address_parts.append(picking.partner_id.street2)
                if picking.partner_id.city:
                    address_parts.append(picking.partner_id.city)
                if picking.partner_id.state_id:
                    address_parts.append(picking.partner_id.state_id.name)
                if picking.partner_id.zip:
                    address_parts.append(picking.partner_id.zip)
                if picking.partner_id.country_id:
                    address_parts.append(picking.partner_id.country_id.name)
                delivery_address = ', '.join(address_parts)
                district = picking.partner_id.city or ''
                state_name = picking.partner_id.state_id.name if picking.partner_id.state_id else ''

            # Get warehouse location from warehouse_id
            warehouse_location = ''
            if picking.picking_type_id and picking.picking_type_id.warehouse_id:
                warehouse = picking.picking_type_id.warehouse_id
                if warehouse.partner_id:
                    wh_parts = [warehouse.name]
                    if warehouse.partner_id.street:
                        wh_parts.append(warehouse.partner_id.street)
                    if warehouse.partner_id.city:
                        wh_parts.append(warehouse.partner_id.city)
                    warehouse_location = ', '.join(wh_parts)
                else:
                    warehouse_location = warehouse.name
            
            # Determine courier or direct
            courier_direct = ''
            if hasattr(picking, 'carrier_id') and picking.carrier_id:
                courier_direct = f"Courier: {picking.carrier_id.name}"
            else:
                courier_direct = "Direct"
            
            # Get workorder number from sale_id
            workorder_number = picking.sale_id.name if picking.sale_id else ''
            
            # Process stock.picking.item.line records
            if hasattr(picking, 'item_line_ids') and picking.item_line_ids:
                for item_line in picking.item_line_ids:
                    # Get products linked to this item through stock.move
                    moves = self.env['stock.move'].search([
                        ('picking_id', '=', picking.id),
                        ('state', '!=', 'cancel')
                    ])
                    
                    # Try to filter moves related to this item, but if no match, use all moves
                    item_moves = []
                    if hasattr(item_line, 'item_details_id') and item_line.item_details_id:
                        item_moves = moves.filtered(lambda m: 
                            hasattr(m.product_id, 'item_details_id') and 
                            m.product_id.item_details_id == item_line.item_details_id
                        )
                    
                    # If no specific item-product link found, use all moves for this picking
                    if not item_moves:
                        item_moves = moves
                    
                    if item_moves:
                        for move in item_moves:
                            # Get serial numbers from stock.move.line
                            move_lines = self.env['stock.move.line'].search([
                                ('move_id', '=', move.id),
                                ('lot_id', '!=', False)
                            ])
                            
                            if move_lines:
                                for move_line in move_lines:
                                    preview_data.append({
                                        'dc_number': picking.dc_number or picking.name,
                                        'dc_delivery_address': delivery_address,
                                        'dc_district': district,
                                        'dc_state': state_name,
                                        'dc_warehouse_location': warehouse_location,
                                        'dc_courier_direct': courier_direct,
                                        'workorder_number': workorder_number,
                                        'dc_customer_po_number': picking.sale_id.client_order_ref or '' if picking.sale_id else '',
                                        'dc_po_issue_date': picking.sale_id.po_issue_date if picking.sale_id and picking.sale_id.po_issue_date else False,
                                        'dc_item_description': item_line.description or '',
                                        'dc_item_name': item_line.item_name or '',
                                        'dc_individual_price': item_line.price_unit,
                                        'dc_hsn_code': item_line.hsn_code or '',
                                        'dc_quotation_quantity': item_line.order_qty,
                                        'dc_tax_amount': item_line.price_tax,
                                        'dc_total_amount': item_line.price_total,
                                        'dc_product_name': move.product_id.name,
                                        'dc_serial_number': move_line.lot_id.name,
                                    })
                            else:
                                # No serial numbers for this product
                                preview_data.append({
                                    'dc_number': picking.dc_number or picking.name,
                                    'dc_delivery_address': delivery_address,
                                    'dc_district': district,
                                    'dc_state': state_name,
                                    'dc_warehouse_location': warehouse_location,
                                    'dc_courier_direct': courier_direct,
                                    'workorder_number': workorder_number,
                                    'dc_customer_po_number': picking.sale_id.client_order_ref or '' if picking.sale_id else '',
                                    'dc_po_issue_date': picking.sale_id.po_issue_date if picking.sale_id and picking.sale_id.po_issue_date else False,
                                    'dc_item_description': item_line.description or '',
                                    'dc_item_name': item_line.item_name or '',
                                    'dc_individual_price': item_line.price_unit,
                                    'dc_hsn_code': item_line.hsn_code or '',
                                    'dc_quotation_quantity': item_line.order_qty,
                                    'dc_tax_amount': item_line.price_tax,
                                    'dc_total_amount': item_line.price_total,
                                    'dc_product_name': move.product_id.name,
                                    'dc_serial_number': 'No Serial Number',
                                })
                    else:
                        # No related products found
                        preview_data.append({
                            'dc_number': picking.dc_number or picking.name,
                            'dc_delivery_address': delivery_address,
                            'dc_district': district,
                            'dc_state': state_name,
                            'dc_warehouse_location': warehouse_location,
                            'dc_courier_direct': courier_direct,
                            'workorder_number': workorder_number,
                            'dc_customer_po_number': picking.sale_id.client_order_ref or '' if picking.sale_id else '',
                            'dc_po_issue_date': picking.sale_id.po_issue_date if picking.sale_id and picking.sale_id.po_issue_date else False,
                            'dc_item_description': item_line.description or '',
                            'dc_item_name': item_line.item_name or '',
                            'dc_individual_price': item_line.price_unit,
                            'dc_hsn_code': item_line.hsn_code or '',
                            'dc_quotation_quantity': item_line.order_qty,
                            'dc_tax_amount': item_line.price_tax,
                            'dc_total_amount': item_line.price_total,
                            'dc_product_name': '',
                            'dc_serial_number': 'No Products Found',
                        })
            else:
                # No item lines found, create basic row
                preview_data.append({
                    'dc_number': picking.dc_number or picking.name,
                    'dc_delivery_address': delivery_address,
                    'dc_district': district,
                    'dc_state': state_name,
                    'dc_warehouse_location': warehouse_location,
                    'dc_courier_direct': courier_direct,
                    'workorder_number': workorder_number,
                    'dc_customer_po_number': picking.sale_id.client_order_ref or '' if picking.sale_id else '',
                    'dc_po_issue_date': picking.sale_id.po_issue_date if picking.sale_id and picking.sale_id.po_issue_date else False,
                    'dc_item_description': '',
                    'dc_item_name': '',
                    'dc_individual_price': 0,
                    'dc_hsn_code': '',
                    'dc_quotation_quantity': 0,
                    'dc_tax_amount': 0,
                    'dc_total_amount': 0,
                    'dc_product_name': '',
                    'dc_serial_number': 'No Items Found',
                })
        
        return preview_data
    
    def _get_enquiry_preview_data(self):
        """Get preview data for enquiry report"""
        domain = []
        
        # Build domain filters
        if self.enquiry_from_date:
            from_datetime = fields.Datetime.to_datetime(self.enquiry_from_date)
            domain.append(('create_date', '>=', from_datetime))
        
        if self.enquiry_to_date:
            to_datetime = fields.Datetime.to_datetime(self.enquiry_to_date).replace(hour=23, minute=59, second=59)
            domain.append(('create_date', '<=', to_datetime))
        
        if self.enquiry_customer_name:
            domain.append('|')
            domain.append(('company', 'ilike', self.enquiry_customer_name))
            domain.append(('company_id.name', 'ilike', self.enquiry_customer_name))
        
        if self.enquiry_status:
            domain.append(('status', '=', self.enquiry_status))
        
        enquiries = self.env['customer.enq'].search(domain, order='create_date asc')
        
        preview_data = []
        
        for enquiry in enquiries:
            # Get contact person name safely
            contact_person = ''
            if hasattr(enquiry, 'customer_name') and enquiry.customer_name:
                if hasattr(enquiry.customer_name, 'name'):
                    contact_person = enquiry.customer_name.name
                else:
                    contact_person = str(enquiry.customer_name)
            
            # Build customer address
            customer_address = ''
            if hasattr(enquiry, 'customer_name') and enquiry.customer_name:
                address_parts = []
                if hasattr(enquiry, 'street') and enquiry.street:
                    address_parts.append(enquiry.street)
                if hasattr(enquiry, 'street2') and enquiry.street2:
                    address_parts.append(enquiry.street2)
                if hasattr(enquiry, 'city') and enquiry.city:
                    address_parts.append(enquiry.city)
                if hasattr(enquiry, 'state_id') and enquiry.state_id:
                    address_parts.append(enquiry.state_id.name)
                if hasattr(enquiry, 'zip') and enquiry.zip:
                    address_parts.append(enquiry.zip)
                if hasattr(enquiry, 'country_id') and enquiry.country_id:
                    address_parts.append(enquiry.country_id.name)
                customer_address = ', '.join(address_parts) if address_parts else ''
            
            # Get customer type display name
            customer_type = ''
            if hasattr(enquiry, 'customer_type') and enquiry.customer_type:
                customer_type_dict = dict(enquiry._fields['customer_type'].selection)
                customer_type = customer_type_dict.get(enquiry.customer_type, enquiry.customer_type)
            
            # Get status display name
            status = ''
            if hasattr(enquiry, 'status') and enquiry.status:
                status_dict = dict(enquiry._fields['status'].selection)
                status = status_dict.get(enquiry.status, enquiry.status)
            
            # Get company (new) - for new customers
            company_new = ''
            if hasattr(enquiry, 'company_partner') and enquiry.company_partner:
                if hasattr(enquiry.company_partner, 'name'):
                    company_new = enquiry.company_partner.name
                else:
                    company_new = str(enquiry.company_partner)
            
            # Get company (existing) - for existing customers  
            company_existing = ''
            if hasattr(enquiry, 'customer_name') and enquiry.customer_name:
                if hasattr(enquiry.customer_name, 'name'):
                    company_existing = enquiry.customer_name.name
                else:
                    company_existing = str(enquiry.customer_name)
            
            preview_data.append({
                'enquiry_number': enquiry.name or '',
                'enquiry_date': enquiry.create_date.date() if enquiry.create_date else '',
                'customer_name': f"{customer_type} | {company_new or company_existing}",
                'contact_person': contact_person,
                'address_full': customer_address,
                'phone_number': getattr(enquiry, 'phone_number', '') or '',
                'email': getattr(enquiry, 'email', '') or '',
                'gst_number': getattr(enquiry, 'gst_number', '') or '',
                'sales_person': enquiry.salesperson_id.name if enquiry.salesperson_id else '',
                'status': f"{status} | Source: {getattr(enquiry, 'source', '') or 'N/A'}",
            })
        
        return preview_data


    def generate_dc_report(self):
        """Generate DC Report Excel file"""
        domain = []
        
        # Build domain filters (same as preview)
        domain.append(('picking_type_id.code', '=', 'outgoing'))
        domain.append(('state', '!=', 'cancel'))
        
        # if self.dc_number:
        #     domain.append(('dc_number', 'ilike', self.dc_number))

        if self.dc_number_ids:
            domain.append(('id', 'in', self.dc_number_ids.ids))

        if self.dc_workorder_number_ids:
            domain.append(('sale_id', 'in', self.dc_workorder_number_ids.ids))
        
        if self.dc_from_date:
            domain.append(('scheduled_date', '>=', self.dc_from_date))
        
        if self.dc_to_date:
            domain.append(('scheduled_date', '<=', self.dc_to_date))
        
        if self.dc_customer_id:
            domain.append(('partner_id', '=', self.dc_customer_id.id))
        
        # if self.dc_workorder_number:
        #     domain.append(('sale_id.name', 'ilike', self.dc_workorder_number))
        
        stock_pickings = self.env['stock.picking'].search(domain, order='scheduled_date asc')
        
        # Create Excel file
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('DC Report')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'valign': 'vcenter'
        })
        
        # Headers
        headers = [
            'DC Number', 'DC Status', 'DC Done Date', 'Delivery Address', 'District', 'State',
            'Warehouse Location', 'Courier/Direct',
            'WorkOrder Number', 'Customer PO Number', 'PO Issue Date', 'Item Name', 'Item Description', 'Product Name',
            'Unit Price', 'HSN Code', 'Quantity', 'Taxes', 'SGST Amount', 'CGST Amount',
            'IGST Amount', 'Tax Amount', 'Tax Excluded', 'Total Amount',
            'Serial Number'
        ]

        # Set column widths
        widths = [20, 18, 15, 40, 20, 20, 25, 15, 20, 22, 15, 25, 40, 30, 15, 15, 12, 20, 15, 15, 15, 15, 15, 15, 20]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)

        # Write headers
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        # row = 1
        # for picking in stock_pickings:
        #     # Get delivery address
        #     delivery_address = ''
        #     if picking.partner_id:
        #         address_parts = []
        #         if picking.partner_id.street:
        #             address_parts.append(picking.partner_id.street)
        #         if picking.partner_id.street2:
        #             address_parts.append(picking.partner_id.street2)
        #         if picking.partner_id.city:
        #             address_parts.append(picking.partner_id.city)
        #         if picking.partner_id.state_id:
        #             address_parts.append(picking.partner_id.state_id.name)
        #         if picking.partner_id.zip:
        #             address_parts.append(picking.partner_id.zip)
        #         if picking.partner_id.country_id:
        #             address_parts.append(picking.partner_id.country_id.name)
        #         delivery_address = ', '.join(address_parts)
            
        #     # Get warehouse location
        #     warehouse_location = ''
        #     if picking.picking_type_id and picking.picking_type_id.warehouse_id:
        #         warehouse = picking.picking_type_id.warehouse_id
        #         if warehouse.partner_id:
        #             wh_parts = [warehouse.name]
        #             if warehouse.partner_id.street:
        #                 wh_parts.append(warehouse.partner_id.street)
        #             if warehouse.partner_id.city:
        #                 wh_parts.append(warehouse.partner_id.city)
        #             warehouse_location = ', '.join(wh_parts)
        #         else:
        #             warehouse_location = warehouse.name
            
        #     # Determine courier or direct
        #     courier_direct = ''
        #     if hasattr(picking, 'carrier_id') and picking.carrier_id:
        #         courier_direct = f"Courier: {picking.carrier_id.name}"
        #     else:
        #         courier_direct = "Direct"
            
        #     # Get workorder number
        #     workorder_number = picking.sale_id.name if picking.sale_id else ''
            
        #     # Process stock.picking.item.line records
        #     if hasattr(picking, 'item_line_ids') and picking.item_line_ids:
        #         for item_line in picking.item_line_ids:
        #             # Get products linked to this item through stock.move
        #             moves = self.env['stock.move'].search([
        #                 ('picking_id', '=', picking.id),
        #                 ('state', '!=', 'cancel')
        #             ])
                    
        #             # Try to filter moves related to this item, but if no match, use all moves
        #             item_moves = []
        #             if hasattr(item_line, 'item_details_id') and item_line.item_details_id:
        #                 item_moves = moves.filtered(lambda m: 
        #                     hasattr(m.product_id, 'item_details_id') and 
        #                     m.product_id.item_details_id == item_line.item_details_id
        #                 )
                    
        #             # If no specific item-product link found, use all moves for this picking
        #             if not item_moves:
        #                 item_moves = moves
                    
        #             if item_moves:
        #                 for move in item_moves:
        #                     # Get serial numbers from stock.move.line
        #                     move_lines = self.env['stock.move.line'].search([
        #                         ('move_id', '=', move.id),
        #                         ('lot_id', '!=', False)
        #                     ])
                            
        #                     if move_lines:
        #                         for move_line in move_lines:
        #                             data = [
        #                                 picking.dc_number or picking.name,
        #                                 delivery_address,
        #                                 warehouse_location,
        #                                 courier_direct,
        #                                 workorder_number,
        #                                 item_line.description or '',
        #                                 item_line.item_name or '',
        #                                 item_line.price_unit,
        #                                 item_line.hsn_code or '',
        #                                 item_line.order_qty,
        #                                 item_line.price_tax,
        #                                 item_line.price_total,
        #                                 move.product_id.name,
        #                                 move_line.lot_id.name,
        #                             ]
                                    
        #                             # Write data with appropriate formatting
        #                             for col, val in enumerate(data):
        #                                 fmt = cell_format
        #                                 if col in [7, 9, 10, 11]:  # Price, qty, tax, total columns
        #                                     fmt = number_format
        #                                 worksheet.write(row, col, val, fmt)
                                    
        #                             row += 1
        #                     else:
        #                         # No serial numbers
        #                         data = [
        #                             picking.dc_number or picking.name,
        #                             delivery_address,
        #                             warehouse_location,
        #                             courier_direct,
        #                             workorder_number,
        #                             item_line.description or '',
        #                             item_line.item_name or '',
        #                             item_line.price_unit,
        #                             item_line.hsn_code or '',
        #                             item_line.order_qty,
        #                             item_line.price_tax,
        #                             item_line.price_total,
        #                             move.product_id.name,
        #                             'No Serial Number',
        #                         ]
                                
        #                         for col, val in enumerate(data):
        #                             fmt = cell_format
        #                             if col in [7, 9, 10, 11]:
        #                                 fmt = number_format
        #                             worksheet.write(row, col, val, fmt)
                                
        #                         row += 1
        #             else:
        #                 # No related products found
        #                 data = [
        #                     picking.dc_number or picking.name,
        #                     delivery_address,
        #                     warehouse_location,
        #                     courier_direct,
        #                     workorder_number,
        #                     item_line.description or '',
        #                     item_line.item_name or '',
        #                     item_line.price_unit,
        #                     item_line.hsn_code or '',
        #                     item_line.order_qty,
        #                     item_line.price_tax,
        #                     item_line.price_total,
        #                     '',
        #                     'No Products Found',
        #                 ]
                        
        #                 for col, val in enumerate(data):
        #                     fmt = cell_format
        #                     if col in [7, 9, 10, 11]:
        #                         fmt = number_format
        #                     worksheet.write(row, col, val, fmt)
                        
        #                 row += 1
        #     else:
        #         # No item lines found
        #         data = [
        #             picking.dc_number or picking.name,
        #             delivery_address,
        #             warehouse_location,
        #             courier_direct,
        #             workorder_number,
        #             '', '', 0, '', 0, 0, 0, '', 'No Items Found'
        #         ]
                
        #         for col, val in enumerate(data):
        #             fmt = cell_format
        #             if col in [7, 9, 10, 11]:
        #                 fmt = number_format
        #             worksheet.write(row, col, val, fmt)
                
        #         row += 1

        # Write data
        row = 1
        for picking in stock_pickings:
            start_row = row

            # Get delivery address
            delivery_address = ''
            district = ''
            state_name = ''
            if picking.partner_id:
                address_parts = []
                if picking.partner_id.street:
                    address_parts.append(picking.partner_id.street)
                if picking.partner_id.street2:
                    address_parts.append(picking.partner_id.street2)
                if picking.partner_id.city:
                    address_parts.append(picking.partner_id.city)
                if picking.partner_id.state_id:
                    address_parts.append(picking.partner_id.state_id.name)
                if picking.partner_id.zip:
                    address_parts.append(picking.partner_id.zip)
                if picking.partner_id.country_id:
                    address_parts.append(picking.partner_id.country_id.name)
                delivery_address = ', '.join(address_parts)
                district = picking.partner_id.city or ''
                state_name = picking.partner_id.state_id.name if picking.partner_id.state_id else ''

            # Get warehouse location
            warehouse_location = ''
            if picking.picking_type_id and picking.picking_type_id.warehouse_id:
                warehouse = picking.picking_type_id.warehouse_id
                if warehouse.partner_id:
                    wh_parts = [warehouse.name]
                    if warehouse.partner_id.street:
                        wh_parts.append(warehouse.partner_id.street)
                    if warehouse.partner_id.city:
                        wh_parts.append(warehouse.partner_id.city)
                    warehouse_location = ', '.join(wh_parts)
                else:
                    warehouse_location = warehouse.name

            # Courier or direct
            courier_direct = ''
            if hasattr(picking, 'carrier_id') and picking.carrier_id:
                courier_direct = f"Courier: {picking.carrier_id.name}"
            else:
                courier_direct = "Direct"

            workorder_number = picking.sale_id.name if picking.sale_id else ''
            dc_po_issue_date = picking.sale_id.po_issue_date if picking.sale_id and picking.sale_id.po_issue_date else ''

            _dc_status_map = {
                'draft': 'Draft', 'waiting': 'Waiting Another Operation', 'confirmed': 'Waiting',
                'assigned': 'Inspection', 'outward': 'Outward', 'inward': 'Inward',
                'done': 'Done', 'cancel': 'Cancelled',
            }
            dc_status = _dc_status_map.get(picking.state, picking.state or '')
            dc_done_date = picking.date_done.strftime('%d/%m/%Y') if picking.date_done else ''

            if hasattr(picking, 'item_line_ids') and picking.item_line_ids:
                for item_line in picking.item_line_ids:
                    moves = self.env['stock.move'].search([
                        ('picking_id', '=', picking.id),
                        ('state', '!=', 'cancel')
                    ])

                    item_moves = []
                    if hasattr(item_line, 'item_details_id') and item_line.item_details_id:
                        item_moves = moves.filtered(lambda m:
                            hasattr(m.product_id, 'item_details_id') and
                            m.product_id.item_details_id == item_line.item_details_id
                        )

                    if not item_moves:
                        item_moves = moves

                    if item_moves:
                        for move in item_moves:
                            move_lines = self.env['stock.move.line'].search([
                                ('move_id', '=', move.id),
                                ('lot_id', '!=', False)
                            ])

                            taxes_str = ', '.join(item_line.tax_id.mapped('name')) if hasattr(item_line, 'tax_id') and item_line.tax_id else ''
                            igst = item_line.igst_amount if hasattr(item_line, 'igst_amount') and item_line.igst_amount else 0.0
                            if igst:
                                sgst = 0.0
                                cgst = 0.0
                            else:
                                sgst = item_line.sgst_amount if hasattr(item_line, 'sgst_amount') and item_line.sgst_amount else 0.0
                                cgst = item_line.cgst_amount if hasattr(item_line, 'cgst_amount') and item_line.cgst_amount else 0.0

                            customer_po = picking.sale_id.client_order_ref or '' if picking.sale_id else ''
                            if move_lines:
                                for move_line in move_lines:
                                    data = [
                                        '',  # Column A - filled by merge below
                                        '',  # Column B - DC Status, filled by merge below
                                        '',  # Column C - DC Done Date, filled by merge below
                                        delivery_address,
                                        district,
                                        state_name,
                                        warehouse_location,
                                        courier_direct,
                                        workorder_number,
                                        customer_po,
                                        dc_po_issue_date,
                                        item_line.item_name or '',
                                        item_line.description or '',
                                        move.product_id.name,
                                        item_line.price_unit,
                                        item_line.hsn_code or '',
                                        item_line.order_qty,
                                        taxes_str,
                                        sgst,
                                        cgst,
                                        igst,
                                        sgst + cgst + igst,
                                        item_line.price_subtotal or 0.0,
                                        item_line.price_total,
                                        move_line.lot_id.name,
                                    ]
                                    for col, val in enumerate(data):
                                        fmt = cell_format
                                        if col == 10:
                                            fmt = date_format
                                        elif col in [14, 16, 18, 19, 20, 21, 22, 23]:
                                            fmt = number_format
                                        worksheet.write(row, col, val, fmt)
                                    row += 1
                            else:
                                data = [
                                    '', '', '',
                                    delivery_address,
                                    district,
                                    state_name,
                                    warehouse_location,
                                    courier_direct,
                                    workorder_number,
                                    customer_po,
                                    dc_po_issue_date,
                                    item_line.item_name or '',
                                    item_line.description or '',
                                    move.product_id.name,
                                    item_line.price_unit,
                                    item_line.hsn_code or '',
                                    item_line.order_qty,
                                    taxes_str,
                                    sgst,
                                    cgst,
                                    igst,
                                    sgst + cgst + igst,
                                    item_line.price_subtotal or 0.0,
                                    item_line.price_total,
                                    'No Serial Number',
                                ]
                                for col, val in enumerate(data):
                                    fmt = cell_format
                                    if col == 10:
                                        fmt = date_format
                                    elif col in [14, 16, 18, 19, 20, 21, 22, 23]:
                                        fmt = number_format
                                    worksheet.write(row, col, val, fmt)
                                row += 1
                    else:
                        customer_po = picking.sale_id.client_order_ref or '' if picking.sale_id else ''
                        taxes_str = ', '.join(item_line.tax_id.mapped('name')) if hasattr(item_line, 'tax_id') and item_line.tax_id else ''
                        igst = item_line.igst_amount if hasattr(item_line, 'igst_amount') and item_line.igst_amount else 0.0
                        if igst:
                            sgst = 0.0
                            cgst = 0.0
                        else:
                            sgst = item_line.sgst_amount if hasattr(item_line, 'sgst_amount') and item_line.sgst_amount else 0.0
                            cgst = item_line.cgst_amount if hasattr(item_line, 'cgst_amount') and item_line.cgst_amount else 0.0
                        data = [
                            '', '', '',
                            delivery_address,
                            district,
                            state_name,
                            warehouse_location,
                            courier_direct,
                            workorder_number,
                            customer_po,
                            dc_po_issue_date,
                            item_line.item_name or '',
                            item_line.description or '',
                            '',
                            item_line.price_unit,
                            item_line.hsn_code or '',
                            item_line.order_qty,
                            taxes_str,
                            sgst,
                            cgst,
                            igst,
                            sgst + cgst + igst,
                            item_line.price_subtotal or 0.0,
                            item_line.price_total,
                            'No Products Found',
                        ]
                        for col, val in enumerate(data):
                            fmt = cell_format
                            if col == 10:
                                fmt = date_format
                            elif col in [14, 16, 18, 19, 20, 21, 22, 23]:
                                fmt = number_format
                            worksheet.write(row, col, val, fmt)
                        row += 1
            else:
                customer_po = picking.sale_id.client_order_ref or '' if picking.sale_id else ''
                data = [
                    '', '', '',
                    delivery_address,
                    district,
                    state_name,
                    warehouse_location,
                    courier_direct,
                    workorder_number,
                    customer_po,
                    dc_po_issue_date,
                    '', '', '', 0.0, '', 0, '', 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 'No Items Found'
                ]
                for col, val in enumerate(data):
                    fmt = cell_format
                    if col == 10:
                        fmt = date_format
                    elif col in [14, 16, 18, 19, 20, 21, 22, 23]:
                        fmt = number_format
                    worksheet.write(row, col, val, fmt)
                row += 1

            end_row = row - 1

            # Merge DC Number, DC Status, DC Done Date columns vertically
            if end_row > start_row:
                worksheet.merge_range(start_row, 0, end_row, 0, picking.dc_number or picking.name or '', cell_format)
                worksheet.merge_range(start_row, 1, end_row, 1, dc_status, cell_format)
                worksheet.merge_range(start_row, 2, end_row, 2, dc_done_date, cell_format)
            else:
                worksheet.write(start_row, 0, picking.dc_number or picking.name or '', cell_format)
                worksheet.write(start_row, 1, dc_status, cell_format)
                worksheet.write(start_row, 2, dc_done_date, cell_format)
        
        workbook.close()
        output.seek(0)
        
        # Save file
        filter_parts = []
        if self.dc_from_date and self.dc_to_date:
            filter_parts.append(f'{self.dc_from_date}_{self.dc_to_date}')
        if self.dc_customer_id:
            filter_parts.append(f'customer_{self.dc_customer_id.name.replace(" ", "_")}')
        
        filename = f'dc_report_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename
        })
        
        self.generate_preview()
        
        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def generate_enquiry_report(self):
        """Generate Enquiry Report Excel file"""
        domain = []
        
        # Build domain filters (same as preview)
        if self.enquiry_from_date:
            from_datetime = fields.Datetime.to_datetime(self.enquiry_from_date)
            domain.append(('create_date', '>=', from_datetime))
        
        if self.enquiry_to_date:
            to_datetime = fields.Datetime.to_datetime(self.enquiry_to_date).replace(hour=23, minute=59, second=59)
            domain.append(('create_date', '<=', to_datetime))
        
        if self.enquiry_customer_name:
            domain.append('|')
            domain.append(('company', 'ilike', self.enquiry_customer_name))
            domain.append(('company_id.name', 'ilike', self.enquiry_customer_name))
        
        if self.enquiry_status:
            domain.append(('status', '=', self.enquiry_status))
        
        enquiries = self.env['customer.enq'].search(domain, order='create_date asc')
        
        # Create Excel file
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Enquiry Report')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'valign': 'vcenter'
        })
        
        # Headers
        headers = [
            'Enquiry Number', 'Enquiry Date', 'Customer Type', 'Company (New)', 
            'Company (Existing)', 'Contact Person', 'Customer Address', 'Phone Number',
            'Email', 'GST Number', 'Salesperson', 'Status'
        ]
        
        # Set column widths
        widths = [20, 15, 18, 30, 30, 25, 45, 18, 30, 20, 20, 18]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
        
        # Write headers
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        row = 1
        for enquiry in enquiries:
            # Get contact person name safely
            contact_person = ''
            if hasattr(enquiry, 'customer_name') and enquiry.customer_name:
                if hasattr(enquiry.customer_name, 'name'):
                    contact_person = enquiry.customer_name.name
                else:
                    contact_person = str(enquiry.customer_name)
            
            # Build customer address
            customer_address = ''
            if hasattr(enquiry, 'customer_name') and enquiry.customer_name:
                address_parts = []
                if hasattr(enquiry, 'street') and enquiry.street:
                    address_parts.append(enquiry.street)
                if hasattr(enquiry, 'street2') and enquiry.street2:
                    address_parts.append(enquiry.street2)
                if hasattr(enquiry, 'city') and enquiry.city:
                    address_parts.append(enquiry.city)
                if hasattr(enquiry, 'state_id') and enquiry.state_id:
                    address_parts.append(enquiry.state_id.name)
                if hasattr(enquiry, 'zip') and enquiry.zip:
                    address_parts.append(enquiry.zip)
                if hasattr(enquiry, 'country_id') and enquiry.country_id:
                    address_parts.append(enquiry.country_id.name)
                customer_address = ', '.join(address_parts) if address_parts else ''
            
            # Get customer type display name
            customer_type = ''
            if hasattr(enquiry, 'customer_type') and enquiry.customer_type:
                customer_type_dict = dict(enquiry._fields['customer_type'].selection)
                customer_type = customer_type_dict.get(enquiry.customer_type, enquiry.customer_type)
            
            # Get status display name
            status = ''
            if hasattr(enquiry, 'status') and enquiry.status:
                status_dict = dict(enquiry._fields['status'].selection)
                status = status_dict.get(enquiry.status, enquiry.status)
            
            # Get company based on customer_type
            company_new = ''
            company_existing = ''

            if hasattr(enquiry, 'customer_type') and enquiry.customer_type:
                if enquiry.customer_type == 'new':
                    # For new customers, use company field (char field)
                    if hasattr(enquiry, 'company') and enquiry.company:
                        company_new = enquiry.company
                else:  # existing customer
                    # For existing customers, use company_id field (Many2one)
                    if hasattr(enquiry, 'company_id') and enquiry.company_id:
                        company_existing = enquiry.company_id.name if hasattr(enquiry.company_id, 'name') else str(enquiry.company_id)
            
            data = [
                enquiry.name or '',
                enquiry.create_date.date() if enquiry.create_date else '',
                customer_type,
                company_new,
                company_existing,
                contact_person,
                customer_address,
                getattr(enquiry, 'phone', '') or '',
                getattr(enquiry, 'email', '') or '',
                getattr(enquiry, 'gst', '') or '',
                enquiry.salesperson_id.name if enquiry.salesperson_id else '',
                status,
            ]
            
            # Write data with appropriate formatting
            for col, val in enumerate(data):
                fmt = cell_format
                if col == 1:  # Enquiry Date
                    fmt = date_format
                worksheet.write(row, col, val, fmt)
            
            row += 1
        
        workbook.close()
        output.seek(0)
        
        # Save file
        filter_parts = []
        if self.enquiry_from_date and self.enquiry_to_date:
            filter_parts.append(f'{self.enquiry_from_date}_{self.enquiry_to_date}')
        if self.enquiry_customer_name:
            filter_parts.append(f'customer_{self.enquiry_customer_name.replace(" ", "_")}')
        if self.enquiry_status:
            filter_parts.append(f'status_{self.enquiry_status}')
        
        filename = f'enquiry_report_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename
        })
        
        self.generate_preview()
        
        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }


    def _get_item_summary_preview_data(self):
        """Get preview data for item summary report based on custom item model"""
        domain = []
        
        # Add item name filter if specified
        if self.item_name:
            domain.append(('item_name', 'ilike', self.item_name))
        
        # Search in your custom item details model (adjust model name as needed)
        # Based on your other reports, it looks like you use 'item.details' model
        items = self.env['item.details'].search(domain, order='item_name asc')
        
        preview_data = []
        
        for item in items:
            # Get stock information from related products
            current_stock = 0.0
            stock_location = ''
            
            # If item has linked products, get their stock
            if hasattr(item, 'product_ids') and item.product_ids:
                for product in item.product_ids:
                    # Get stock quants for this product
                    quants = self.env['stock.quant'].search([
                        ('product_id', '=', product.id),
                        ('location_id.usage', '=', 'internal'),
                        ('quantity', '>', 0)
                    ])
                    
                    if quants:
                        current_stock += sum(quants.mapped('quantity'))
                        if not stock_location and quants:
                            main_location = quants[0].location_id
                            stock_location = main_location.complete_name if main_location else ''
            
            # Get supplier info from linked products
            main_supplier = ''
            sales_price = 0.0
            cost_price = 0.0
            hsn_code = ''
            barcode = ''
            unit_of_measure = ''
            
            if hasattr(item, 'product_ids') and item.product_ids:
                first_product = item.product_ids[0]
                sales_price = first_product.list_price
                cost_price = first_product.standard_price
                
                if first_product.seller_ids:
                    main_supplier = first_product.seller_ids[0].partner_id.name
                
                if hasattr(first_product, 'l10n_in_hsn_code'):
                    hsn_code = first_product.l10n_in_hsn_code or ''
                    
                barcode = first_product.barcode or ''
                unit_of_measure = first_product.uom_id.name if first_product.uom_id else ''
            
            # Use item's own pricing if available (adjust field names as per your model)
            if hasattr(item, 'price_unit') and item.price_unit:
                sales_price = item.price_unit
            if hasattr(item, 'cost_price') and item.cost_price:
                cost_price = item.cost_price
            if hasattr(item, 'hsn_code') and item.hsn_code:
                hsn_code = item.hsn_code
                
            # Get category display label from selection field
            category_label = ''
            if hasattr(item, 'category') and item.category:
                category_dict = dict(item._fields['category'].selection)
                category_label = category_dict.get(item.category, item.category)

            preview_data.append({
                'item_summary_name': item.item_name or '',
                'item_summary_code': item.name or '',
                'item_summary_description': getattr(item, 'description', '') or getattr(item, 'item_description', '') or '',
                'item_summary_category': category_label,
                'item_summary_hsn_code': item.hsn_code or '',
            })
        
        return preview_data

    def generate_item_summary_report(self):
        """Generate Item Summary Report Excel file"""
        domain = []
        
        # Add item name filter if specified
        if self.item_name:
            domain.append(('item_name', 'ilike', self.item_name))
        
        # Search in your custom item details model
        items = self.env['item.details'].search(domain, order='item_name asc')
        
        # Create Excel file
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Item Summary')
        
        # Define formats
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter'
        })
        
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter'
        })
        
        quantity_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter',
            'align': 'center'
        })
        
        # Headers
        headers = [
            'Item Code', 'Item Name', 'Description', 'Category', 'HSN Code', 'Products'
        ]
        
        # Set column widths
        widths = [30, 20, 40, 20, 15, 40]
        for idx, width in enumerate(widths):
            worksheet.set_column(idx, idx, width)
        
        # Write headers
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)
        
        # Write data
        # row = 1
        # for item in items:
        #     # Get stock information from related products
        #     current_stock = 0.0
        #     stock_location = ''
            
        #     if hasattr(item, 'product_ids') and item.product_ids:
        #         for product in item.product_ids:
        #             quants = self.env['stock.quant'].search([
        #                 ('product_id', '=', product.id),
        #                 ('location_id.usage', '=', 'internal'),
        #                 ('quantity', '>', 0)
        #             ])
                    
        #             if quants:
        #                 current_stock += sum(quants.mapped('quantity'))
        #                 if not stock_location and quants:
        #                     main_location = quants[0].location_id
        #                     stock_location = main_location.complete_name if main_location else ''
            
        #     # Get supplier and pricing info
        #     main_supplier = ''
        #     sales_price = 0.0
        #     cost_price = 0.0
        #     hsn_code = ''
        #     barcode = ''
        #     unit_of_measure = ''
            
        #     if hasattr(item, 'product_ids') and item.product_ids:
        #         first_product = item.product_ids[0]
        #         sales_price = first_product.list_price
        #         cost_price = first_product.standard_price
                
        #         if first_product.seller_ids:
        #             main_supplier = first_product.seller_ids[0].partner_id.name
                
        #         if hasattr(first_product, 'l10n_in_hsn_code'):
        #             hsn_code = first_product.l10n_in_hsn_code or ''
                    
        #         barcode = first_product.barcode or ''
        #         unit_of_measure = first_product.uom_id.name if first_product.uom_id else ''
            
        #     # Use item's own pricing if available
        #     if hasattr(item, 'price_unit') and item.price_unit:
        #         sales_price = item.price_unit
        #     if hasattr(item, 'cost_price') and item.cost_price:
        #         cost_price = item.cost_price
        #     if hasattr(item, 'hsn_code') and item.hsn_code:
        #         hsn_code = item.hsn_code
                
        #     category_label = ''
        #     if hasattr(item, 'category') and item.category:
        #         category_dict = dict(item._fields['category'].selection)
        #         category_label = category_dict.get(item.category, item.category)

        #     # Get products from item lines
        #     product_names = ''
        #     try:
        #         # Try common field names for the Many2one back-reference
        #         lines = None
        #         for field_name in ['item_id', 'item_details_id', 'details_id', 'parent_id']:
        #             try:
        #                 lines = self.env['item.details.line'].search([
        #                     (field_name, '=', item.id)
        #                 ])
        #                 if lines:
        #                     break
        #             except Exception:
        #                 continue
                
        #         if lines:
        #             product_names = ', '.join(
        #                 l.product_id.name for l in lines if l.product_id
        #             )
        #     except Exception:
        #         product_names = ''

        #     data = [
        #         item.name or '',
        #         item.item_name or '',
        #         getattr(item, 'description', '') or getattr(item, 'item_description', '') or '',
        #         category_label,
        #         item.hsn_code or '',
        #         product_names,
        #     ]

        #     for col, val in enumerate(data):
        #         worksheet.write(row, col, val, cell_format)

        #     row += 1

        row = 1
        for item in items:
            start_row = row

            category_label = ''
            if hasattr(item, 'category') and item.category:
                category_dict = dict(item._fields['category'].selection)
                category_label = category_dict.get(item.category, item.category)

            # Get product lines
            product_lines = []
            try:
                for field_name in ['item_id', 'item_details_id', 'details_id', 'parent_id']:
                    try:
                        lines = self.env['item.details.line'].search([
                            (field_name, '=', item.id)
                        ])
                        if lines:
                            product_lines = lines
                            break
                    except Exception:
                        continue
            except Exception:
                product_lines = []

            common = [
                '',  # Column A - item code, filled by merge below
                item.item_name or '',
                getattr(item, 'description', '') or getattr(item, 'item_description', '') or '',
                category_label,
                item.hsn_code or '',
            ]

            if product_lines:
                for line in product_lines:
                    product_name = line.product_id.name if line.product_id else ''
                    data = common + [product_name]
                    for col, val in enumerate(data):
                        worksheet.write(row, col, val, cell_format)
                    row += 1
            else:
                data = common + ['(No products)']
                for col, val in enumerate(data):
                    worksheet.write(row, col, val, cell_format)
                row += 1

            end_row = row - 1

            # Merge Item Code column (column 0) vertically
            if end_row > start_row:
                worksheet.merge_range(start_row, 0, end_row, 0, item.name or '', cell_format)
            else:
                worksheet.write(start_row, 0, item.name or '', cell_format)
            
        
        workbook.close()
        output.seek(0)
        
        # Save file
        filter_parts = []
        if self.item_summary_from_date and self.item_summary_to_date:
            filter_parts.append(f'{self.item_summary_from_date}_{self.item_summary_to_date}')
        if self.item_name:
            filter_parts.append(f'item_{self.item_name.replace(" ", "_")}')
        
        filename = f'item_summary_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename
        })
        
        self.generate_preview()
        
        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_task_allocation_preview_data(self):
        """Get preview data for Task Allocation Report"""
        domain = []

        if self.task_from_date:
            domain.append(('create_date', '>=', self.task_from_date))

        if self.task_to_date:
            domain.append(('create_date', '<=', self.task_to_date))

        if self.task_employee_id:
            # Match via user linked to employee
            user = self.task_employee_id.user_id
            if user:
                domain.append(('user_ids', 'in', [user.id]))

        if self.task_project_id:
            domain.append(('project_id', '=', self.task_project_id.id))

        if self.task_status:
            domain.append(('stage_id', '=', self.task_stage_id.id))

        tasks = self.env['project.task'].search(domain, order='create_date asc')

        preview_data = []

        for task in tasks:
            # Get assigned employees (can be multiple in Odoo 16+)
            assigned_employees = ''
            if task.user_ids:
                assigned_employees = ', '.join(task.user_ids.mapped('name'))
            elif hasattr(task, 'user_id') and task.user_id:
                assigned_employees = task.user_id.name

            # Get workorder number from project's customer_po or sale_order link
            # Workorder number
            workorder_number = ''
            sale_order_ref = False
            if hasattr(task.project_id, 'customer_po') and task.project_id.customer_po:
                cp = task.project_id.customer_po
                # customer_po is a Many2one to sale.order, get its name
                if hasattr(cp, 'name'):
                    workorder_number = cp.name or ''
                else:
                    workorder_number = str(cp)
                sale_order_ref = cp
            elif hasattr(task, 'sale_line_id') and task.sale_line_id:
                workorder_number = task.sale_line_id.order_id.name or ''
                sale_order_ref = task.sale_line_id.order_id

            # Customer PO Number and PO Issue Date from the linked sale order
            task_customer_po = sale_order_ref.client_order_ref if sale_order_ref and hasattr(sale_order_ref, 'client_order_ref') else ''
            task_po_issue_date = sale_order_ref.po_issue_date if sale_order_ref and hasattr(sale_order_ref, 'po_issue_date') and sale_order_ref.po_issue_date else False

            # Get customer from project's partner
            customer = ''
            if task.project_id and task.project_id.partner_id:
                customer = task.project_id.partner_id.name

            # Get hours logged (timesheet hours)
            hours_logged = 0.0
            if hasattr(task, 'effective_hours'):
                hours_logged = task.effective_hours or 0.0
            elif hasattr(task, 'timesheet_ids'):
                hours_logged = sum(task.timesheet_ids.mapped('unit_amount'))

            # Get task status display label
            status = ''
            if task.state:
                state_dict = dict(task._fields['state'].selection)
                status = task.stage_id.name if task.stage_id else ''
            elif task.stage_id:
                status = task.stage_id.name

            preview_data.append({
                'task_name': task.name or '',
                'task_project': task.project_id.name if task.project_id else '',
                'task_assigned_employee': assigned_employees,
                'task_status': status,
                'task_workorder_number': workorder_number,
                'task_customer_po_number': task_customer_po,
                'task_po_issue_date': task_po_issue_date,
                'task_customer': customer,
                'task_hours_logged': hours_logged,
            })

        return preview_data


    def generate_task_allocation_report(self):
        """Generate Task Allocation Report Excel file"""
        domain = []

        if self.task_from_date:
            domain.append(('create_date', '>=', self.task_from_date))

        if self.task_to_date:
            domain.append(('create_date', '<=', self.task_to_date))

        if self.task_employee_id:
            user = self.task_employee_id.user_id
            if user:
                domain.append(('user_ids', 'in', [user.id]))

        if self.task_project_id:
            domain.append(('project_id', '=', self.task_project_id.id))

        if self.task_status:
            domain.append(('stage_id', '=', self.task_stage_id.id))

        tasks = self.env['project.task'].search(domain, order='create_date asc')

        # ── Excel setup ────────────────────────────────────────────────────────
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Task Allocation Report')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'border': 1,
            'valign': 'vcenter',
        })

        # ── Columns ─────────────────────────────────────────────────────────────
        columns = [
            ('Task Name',          35),
            ('Project Name',       30),
            ('Assigned Employee',  25),
            ('Status',             20),
            ('Workorder Number',   22),
            ('Customer PO Number', 22),
            ('PO Issue Date',      15),
            ('Customer',           25),
            # ('Hours Logged',       15),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)

        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        # ── Data rows ────────────────────────────────────────────────────────────
        row = 1

        for task in tasks:
            # Assigned employees
            assigned_employees = ''
            if task.user_ids:
                assigned_employees = ', '.join(task.user_ids.mapped('name'))
            elif hasattr(task, 'user_id') and task.user_id:
                assigned_employees = task.user_id.name

            # Workorder number
            workorder_number = ''
            task_sale_order = False
            if hasattr(task.project_id, 'customer_po') and task.project_id.customer_po:
                cp = task.project_id.customer_po
                # customer_po is a Many2one to sale.order, get its name
                if hasattr(cp, 'name'):
                    workorder_number = cp.name or ''
                else:
                    workorder_number = str(cp)
                task_sale_order = cp
            elif hasattr(task, 'sale_line_id') and task.sale_line_id:
                workorder_number = task.sale_line_id.order_id.name or ''
                task_sale_order = task.sale_line_id.order_id

            # Customer PO Number and PO Issue Date
            task_customer_po = task_sale_order.client_order_ref if task_sale_order and hasattr(task_sale_order, 'client_order_ref') else ''
            task_po_issue_date = task_sale_order.po_issue_date if task_sale_order and hasattr(task_sale_order, 'po_issue_date') and task_sale_order.po_issue_date else ''

            # Customer
            customer = ''
            if task.project_id and task.project_id.partner_id:
                customer = task.project_id.partner_id.name

            # Hours logged
            hours_logged = 0.0
            if hasattr(task, 'effective_hours'):
                hours_logged = task.effective_hours or 0.0
            elif hasattr(task, 'timesheet_ids'):
                hours_logged = sum(task.timesheet_ids.mapped('unit_amount'))

            # Status
            status = ''
            if task.stage_id:
                status = task.stage_id.name
            elif hasattr(task, 'kanban_state') and task.kanban_state:
                kanban_dict = dict(task._fields['kanban_state'].selection)
                status = kanban_dict.get(task.kanban_state, task.kanban_state)

            data = [
                task.name or '',
                task.project_id.name if task.project_id else '',
                assigned_employees or '',
                status or '',
                workorder_number or '',
                task_customer_po or '',
                task_po_issue_date,
                customer or '',
                # float(hours_logged) if hours_logged else 0.0,   # ensure it's a plain float
            ]

            for col, val in enumerate(data):
                fmt = date_format if col == 6 else cell_format
                worksheet.write(row, col, val, fmt)

            row += 1

        workbook.close()
        output.seek(0)

        # ── Save & return ─────────────────────────────────────────────────────
        filter_parts = []
        if self.task_from_date and self.task_to_date:
            filter_parts.append(f'{self.task_from_date}_{self.task_to_date}')
        if self.task_employee_id:
            filter_parts.append(f'emp_{self.task_employee_id.name.replace(" ", "_")}')
        if self.task_project_id:
            filter_parts.append(f'proj_{self.task_project_id.name.replace(" ", "_")}')

        filename = (
            f'task_allocation_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_installation_summary_preview_data(self):
        """Get preview data for Installation Summary Report"""
        # Get only installation projects (projects linked to installation module)
        installation_projects = self.env['project.project'].search([
            ('name', 'ilike', 'INS')
        ])

        domain = [('project_id', 'in', installation_projects.ids)]

        if self.installation_from_date:
            domain.append(('create_date', '>=', self.installation_from_date))

        if self.installation_to_date:
            domain.append(('create_date', '<=', self.installation_to_date))

        # if self.installation_project_id:
        #     domain.append(('project_id', '=', self.installation_project_id.id))

        if self.installation_project_ids:
            domain.append(('project_id', 'in', self.installation_project_ids.ids))

        if self.installation_customer_id_filter:
            domain.append(('project_id.partner_id', '=', self.installation_customer_id_filter.id))

        tasks = self.env['project.task'].search(domain, order='create_date asc')

        preview_data = []

        for task in tasks:
            # Assigned employees
            assigned_employees = ''
            if task.user_ids:
                assigned_employees = ', '.join(task.user_ids.mapped('name'))
            elif hasattr(task, 'user_id') and task.user_id:
                assigned_employees = task.user_id.name

            # Workorder number
            workorder_number = ''
            if hasattr(task.project_id, 'customer_po') and task.project_id.customer_po:
                cp = task.project_id.customer_po
                workorder_number = cp.name if hasattr(cp, 'name') else str(cp)
            elif hasattr(task, 'sale_line_id') and task.sale_line_id:
                workorder_number = task.sale_line_id.order_id.name or ''

            # Customer
            customer = ''
            if task.project_id and task.project_id.partner_id:
                customer = task.project_id.partner_id.name

            # Hours logged
            hours_logged = 0.0
            if hasattr(task, 'effective_hours'):
                hours_logged = task.effective_hours or 0.0
            elif hasattr(task, 'timesheet_ids'):
                hours_logged = sum(task.timesheet_ids.mapped('unit_amount'))

            # Deadline
            deadline = task.date_deadline if task.date_deadline else False

            customer_po = ''
            inst_po_issue_date = False
            if hasattr(task, 'sale_line_id') and task.sale_line_id and task.sale_line_id.order_id:
                customer_po = task.sale_line_id.order_id.client_order_ref or ''
                inst_po_issue_date = task.sale_line_id.order_id.po_issue_date if task.sale_line_id.order_id.po_issue_date else False
            elif hasattr(task, 'sale_order_id') and task.sale_order_id:
                customer_po = task.sale_order_id.client_order_ref or ''
                inst_po_issue_date = task.sale_order_id.po_issue_date if task.sale_order_id.po_issue_date else False

            # DC data from delivery booking → stock.picking
            dc_number = ''
            dc_tax_excluded = 0.0
            dc_tax_included = 0.0
            dc_status = ''
            if hasattr(task, 'dc_id') and task.dc_id and task.dc_id.dc_id:
                picking = task.dc_id.dc_id
                dc_number = picking.dc_number or ''
                dc_tax_excluded = sum(picking.item_line_ids.mapped('price_subtotal'))
                dc_tax_included = sum(picking.item_line_ids.mapped('price_total'))
                state_labels = dict(self.env['stock.picking']._fields['state'].selection)
                dc_status = state_labels.get(picking.state, picking.state)

            preview_data.append({
                'inst_task_name': task.name or '',
                'inst_project_name': task.project_id.name if task.project_id else '',
                'inst_customer': customer,
                'inst_workorder_number': workorder_number,
                'inst_customer_po_number': customer_po,
                'inst_po_issue_date': inst_po_issue_date,
                'inst_assigned_employee': assigned_employees,
                'inst_deadline': deadline,
                'inst_hours_logged': hours_logged,
                'inst_dc_number': dc_number,
                'inst_dc_tax_excluded': dc_tax_excluded,
                'inst_dc_tax_included': dc_tax_included,
                'inst_dc_status': dc_status,
            })

        return preview_data


    def generate_installation_summary_report(self):
        """Generate Installation Summary Report Excel file"""
        # Get only installation projects (projects linked to installation module)
        installation_projects = self.env['project.project'].search([
            ('name', 'ilike', 'INS')
        ])

        domain = [('project_id', 'in', installation_projects.ids)]

        if self.installation_from_date:
            domain.append(('create_date', '>=', self.installation_from_date))

        if self.installation_to_date:
            domain.append(('create_date', '<=', self.installation_to_date))

        # if self.installation_project_id:
        #     domain.append(('project_id', '=', self.installation_project_id.id))

        if self.installation_project_ids:
            domain.append(('project_id', 'in', self.installation_project_ids.ids))

        if self.installation_customer_id_filter:
            domain.append(('project_id.partner_id', '=', self.installation_customer_id_filter.id))

        tasks = self.env['project.task'].search(domain, order='create_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Installation Summary')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Task Name',               35),  # 0
            ('Project Name',            30),  # 1
            ('Customer',                25),  # 2
            ('Workorder Number',        22),  # 3
            ('Customer PO Number',      22),  # 4
            ('DC Number',               20),  # 5
            ('DC Tax Excluded Value',   20),  # 6
            ('DC Tax Included Value',   20),  # 7
            ('DC Status',               15),  # 8
            ('PO Issue Date',           15),  # 9
            ('Assigned Employee',       25),  # 10
            ('Deadline',                15),  # 11
            ('District',                20),  # 12
            # ('Hours Logged',          15),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        row = 1
        for task in tasks:
            # Assigned employees
            assigned_employees = ''
            if task.user_ids:
                assigned_employees = ', '.join(task.user_ids.mapped('name'))
            elif hasattr(task, 'user_id') and task.user_id:
                assigned_employees = task.user_id.name

            # Workorder number
            workorder_number = ''
            if hasattr(task.project_id, 'customer_po') and task.project_id.customer_po:
                cp = task.project_id.customer_po
                workorder_number = cp.name if hasattr(cp, 'name') else str(cp)
            elif hasattr(task, 'sale_line_id') and task.sale_line_id:
                workorder_number = task.sale_line_id.order_id.name or ''

            # Customer
            customer = ''
            if task.project_id and task.project_id.partner_id:
                customer = task.project_id.partner_id.name

            # Hours logged
            hours_logged = 0.0
            if hasattr(task, 'effective_hours'):
                hours_logged = task.effective_hours or 0.0
            elif hasattr(task, 'timesheet_ids'):
                hours_logged = sum(task.timesheet_ids.mapped('unit_amount'))

            # Customer PO number and PO Issue Date
            customer_po = ''
            inst_po_issue_date = False
            if hasattr(task, 'sale_line_id') and task.sale_line_id and task.sale_line_id.order_id:
                customer_po = task.sale_line_id.order_id.client_order_ref or ''
                inst_po_issue_date = task.sale_line_id.order_id.po_issue_date if task.sale_line_id.order_id.po_issue_date else False
            elif hasattr(task, 'sale_order_id') and task.sale_order_id:
                customer_po = task.sale_order_id.client_order_ref or ''
                inst_po_issue_date = task.sale_order_id.po_issue_date if task.sale_order_id.po_issue_date else False

            # Deadline
            deadline = task.date_deadline if task.date_deadline else False

            # District — city from sale order billing address (partner_invoice_id)
            district = ''
            if task.sale_order_id and task.sale_order_id.partner_invoice_id:
                district = task.sale_order_id.partner_invoice_id.city or ''

            # DC data from delivery booking → stock.picking
            dc_number = ''
            dc_tax_excluded = 0.0
            dc_tax_included = 0.0
            dc_status = ''
            if hasattr(task, 'dc_id') and task.dc_id and task.dc_id.dc_id:
                picking = task.dc_id.dc_id
                dc_number = picking.dc_number or ''
                dc_tax_excluded = sum(picking.item_line_ids.mapped('price_subtotal'))
                dc_tax_included = sum(picking.item_line_ids.mapped('price_total'))
                state_labels = dict(self.env['stock.picking']._fields['state'].selection)
                dc_status = state_labels.get(picking.state, picking.state)

            data = [
                task.name or '',                                   # 0  Task Name
                task.project_id.name if task.project_id else '',   # 1  Project Name
                customer or '',                                    # 2  Customer
                workorder_number or '',                            # 3  Workorder Number
                customer_po,                                       # 4  Customer PO Number
                dc_number,                                         # 5  DC Number
                dc_tax_excluded,                                   # 6  DC Tax Excluded Value
                dc_tax_included,                                   # 7  DC Tax Included Value
                dc_status,                                         # 8  DC Status
                inst_po_issue_date,                                # 9  PO Issue Date
                assigned_employees or '',                          # 10 Assigned Employee
                deadline,                                          # 11 Deadline
                district,                                          # 12 District
                # float(hours_logged) if hours_logged else 0.0,
            ]

            for col, val in enumerate(data):
                if col in {9, 11}:  # PO Issue Date and Deadline
                    fmt = date_format
                elif col in {6, 7}:  # DC Tax Excluded / Included Value
                    fmt = number_format
                else:
                    fmt = cell_format
                worksheet.write(row, col, val, fmt)

            row += 1

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.installation_from_date and self.installation_to_date:
            filter_parts.append(f'{self.installation_from_date}_{self.installation_to_date}')
        # if self.installation_project_id:
        #     filter_parts.append(f'proj_{self.installation_project_id.name.replace(" ", "_")}')
        if self.installation_project_ids:
            for project in self.installation_project_ids:
                filter_parts.append(f'proj_{project.name.replace(" ", "_")}')
        if self.installation_customer_id_filter:
            filter_parts.append(f'cust_{self.installation_customer_id_filter.name.replace(" ", "_")}')

        filename = (
            f'installation_summary_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }
    
    def _get_serial_number_preview_data(self):
        """Get preview data for Serial Number Product Tracking"""
        domain = []

        # if self.serial_product_id:
        #     domain.append(('product_id', '=', self.serial_product_id.id))

        if self.serial_product_ids:
            domain.append(('product_id', 'in', self.serial_product_ids.ids))

        if self.serial_from_date:
            domain.append(('create_date', '>=', self.serial_from_date))

        if self.serial_to_date:
            domain.append(('create_date', '<=', self.serial_to_date))

        lots = self.env['stock.lot'].search(domain, order='create_date asc')

        preview_data = []

        for lot in lots:
            preview_data.append({
                'serial_lot_number': lot.name or '',
                'serial_workorder': lot.sale_order_id.name if lot.sale_order_id else '',
                'serial_product': lot.product_id.name if lot.product_id else '',
                'serial_created_date': lot.create_date.date() if lot.create_date else '',
            })

        return preview_data


    def generate_serial_number_report(self):
        """Generate Serial Number Product Tracking Excel file"""
        domain = []

        # if self.serial_product_id:
        #     domain.append(('product_id', '=', self.serial_product_id.id))

        if self.serial_product_ids:
            domain.append(('product_id', 'in', self.serial_product_ids.ids))

        if self.serial_from_date:
            domain.append(('create_date', '>=', self.serial_from_date))

        if self.serial_to_date:
            domain.append(('create_date', '<=', self.serial_to_date))

        lots = self.env['stock.lot'].search(domain, order='create_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Serial Number Tracking')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Lot/Serial Number', 25),
            ('Workorder',         25),
            ('Product',           35),
            ('Created Date',      18),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        row = 1
        for lot in lots:
            data = [
                lot.name or '',
                lot.sale_order_id.name if lot.sale_order_id else '',
                lot.product_id.name if lot.product_id else '',
                lot.create_date.date() if lot.create_date else '',
            ]

            for col, val in enumerate(data):
                fmt = date_format if col == 3 else cell_format
                worksheet.write(row, col, val, fmt)

            row += 1

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.serial_product_id:
            filter_parts.append(f'product_{self.serial_product_id.name.replace(" ", "_")}')
        if self.serial_from_date and self.serial_to_date:
            filter_parts.append(f'{self.serial_from_date}_{self.serial_to_date}')

        filename = (
            f'serial_number_tracking_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_stock_summary_preview_data(self):
        """Get preview data for Stock Summary Report"""
        part_number_type_labels = {'auto_generate': 'Auto Generate', 'manual': 'Manual'}
        products = self.env['product.product'].search([
            ('detailed_type', '=', 'product'),
            ('qty_available', '>', 0),
        ], order='name asc')
        preview_data = []
        for product in products:
            qty = product.qty_available
            cost = product.list_price
            preview_data.append({
                'stock_product_name': product.name or '',
                'stock_internal_ref': product.default_code or '',
                'stock_category': product.categ_id.name if product.categ_id else '',
                'stock_part_number_type': part_number_type_labels.get(product.product_id_part_number_type, product.product_id_part_number_type or ''),
                'stock_part_number': product.part_number or '',
                'stock_cost_price': cost,
                'stock_qty_on_hand': qty,
                'stock_total_price': qty * cost,
            })
        return preview_data

    def generate_stock_summary_report(self):
        """Generate Stock Summary Report Excel file"""
        part_number_type_labels = {'auto_generate': 'Auto Generate', 'manual': 'Manual'}
        products = self.env['product.product'].search([
            ('detailed_type', '=', 'product'),
            ('qty_available', '>', 0),
        ], order='name asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Stock Summary')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        qty_format = workbook.add_format({
            'num_format': '#,##0', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Product Name',        40),  # 0
            ('Internal Reference',  22),  # 1
            ('Category',            25),  # 2
            ('Part Number Type',    20),  # 3
            ('Part Number',         22),  # 4
            ('Cost Price',          15),  # 5
            ('Qty On Hand',         15),  # 6
            ('Total Price',         18),  # 7
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        row = 1
        for product in products:
            qty = product.qty_available
            cost = product.list_price
            data = [
                product.name or '',
                product.default_code or '',
                product.categ_id.name if product.categ_id else '',
                part_number_type_labels.get(product.product_id_part_number_type, product.product_id_part_number_type or ''),
                product.part_number or '',
                cost,
                qty,
                qty * cost,
            ]
            for col, val in enumerate(data):
                if col in {5, 6, 7}:
                    fmt = qty_format if col == 6 else number_format
                else:
                    fmt = cell_format
                worksheet.write(row, col, val, fmt)
            row += 1

        workbook.close()
        output.seek(0)

        filename = f'stock_summary_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_dc_movement_preview_data(self):
        """Get preview data for DC Movement Report"""
        domain = [('picking_type_code', '=', 'outgoing')]

        # if self.dc_movement_dc_number:
        #     domain.append(('id', '=', self.dc_movement_dc_number.id))

        if self.dc_movement_dc_number_ids:
            domain.append(('id', 'in', self.dc_movement_dc_number_ids.ids))

        if self.dc_movement_from_date:
            domain.append(('create_date_only', '>=', self.dc_movement_from_date))

        if self.dc_movement_to_date:
            domain.append(('create_date_only', '<=', self.dc_movement_to_date))

        if self.dc_movement_customer_id:
            domain.append(('partner_id', '=', self.dc_movement_customer_id.id))

        if self.dc_movement_warehouse_id:
            domain.append(('warehouse_id', '=', self.dc_movement_warehouse_id.id))

        if self.dc_movement_status:
            domain.append(('state', '=', self.dc_movement_status))

        pickings = self.env['stock.picking'].search(domain, order='create_date_only asc')

        preview_data = []

        for picking in pickings:
            state_dict = dict(picking._fields['state'].selection)
            status = state_dict.get(picking.state, picking.state)

            preview_data.append({
                'dc_movement_dc_number_val': picking.dc_number or '',
                'dc_movement_reference': picking.name or '',
                'dc_movement_warehouse': picking.warehouse_id.name if picking.warehouse_id else '',
                'dc_movement_created_date': picking.create_date_only if picking.create_date_only else '',
                'dc_movement_customer': picking.partner_id.name if picking.partner_id else '',
                'dc_movement_source_document': picking.origin or '',
                'dc_movement_status': status,
            })

        return preview_data

    def generate_dc_movement_report(self):
        """Generate DC Movement Report Excel file"""
        domain = [('picking_type_code', '=', 'outgoing')]

        # if self.dc_movement_dc_number:
        #     domain.append(('id', '=', self.dc_movement_dc_number.id))

        if self.dc_movement_dc_number_ids:
            domain.append(('id', 'in', self.dc_movement_dc_number_ids.ids))

        if self.dc_movement_from_date:
            domain.append(('create_date_only', '>=', self.dc_movement_from_date))

        if self.dc_movement_to_date:
            domain.append(('create_date_only', '<=', self.dc_movement_to_date))

        if self.dc_movement_customer_id:
            domain.append(('partner_id', '=', self.dc_movement_customer_id.id))

        if self.dc_movement_warehouse_id:
            domain.append(('warehouse_id', '=', self.dc_movement_warehouse_id.id))

        if self.dc_movement_status:
            domain.append(('state', '=', self.dc_movement_status))

        pickings = self.env['stock.picking'].search(domain, order='create_date_only asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('DC Movement Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('DC Number',       22),
            ('Reference',       22),
            ('Warehouse',       20),
            ('Created Date',    15),
            ('Customer',        30),
            ('Source Document', 22),
            ('Status',          18),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        row = 1
        for picking in pickings:
            state_dict = dict(picking._fields['state'].selection)
            status = state_dict.get(picking.state, picking.state)

            data = [
                picking.dc_number or '',
                picking.name or '',
                picking.warehouse_id.name if picking.warehouse_id else '',
                picking.create_date_only if picking.create_date_only else '',
                picking.partner_id.name if picking.partner_id else '',
                picking.origin or '',
                status,
            ]

            for col, val in enumerate(data):
                fmt = date_format if col == 3 else cell_format
                worksheet.write(row, col, val, fmt)

            row += 1

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.dc_movement_from_date and self.dc_movement_to_date:
            filter_parts.append(f'{self.dc_movement_from_date}_{self.dc_movement_to_date}')
        if self.dc_movement_customer_id:
            filter_parts.append(f'customer_{self.dc_movement_customer_id.name.replace(" ", "_")}')

        filename = (
            f'dc_movement_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }
    
    def _get_courier_preview_data(self):
        """Get preview data for Courier Report"""
        domain = [('dispatch_type', '=', 'courier')]

        if self.courier_from_date:
            domain.append(('courier_booking_date', '>=', self.courier_from_date))

        if self.courier_to_date:
            domain.append(('courier_booking_date', '<=', self.courier_to_date))

        if self.courier_partner_filter_id:
            domain.append(('courier_partner_id', '=', self.courier_partner_filter_id.id))

        if self.courier_order_type:
            domain.append(('order_type', '=', self.courier_order_type))

        bookings = self.env['delivery.booking'].search(domain, order='courier_booking_date asc')

        preview_data = []

        for booking in bookings:
            # order_type_dict = dict(booking._fields['order_type'].selection)
            order_type_map = {
                'elcot': 'ELCOT Order',
                'others': 'Direct Order',
            }
            order_type_label = order_type_map.get(booking.order_type, booking.order_type or '')

            dc_number = ''
            customer = ''
            if booking.dc_id:
                dc_number = booking.dc_id.dc_number or ''
                customer = booking.dc_id.partner_id.name if booking.dc_id.partner_id else ''

            preview_data.append({
                'courier_dc_number': dc_number,
                'courier_sale_number': booking.sale_order_id.name if booking.sale_order_id else '',
                'courier_order_type_val': order_type_label,
                'courier_customer': customer,
                'courier_customer_po_number': booking.sale_order_id.client_order_ref or '' if booking.sale_order_id else '',
                'courier_po_issue_date': booking.sale_order_id.po_issue_date if booking.sale_order_id and booking.sale_order_id.po_issue_date else False,
                'courier_booking_date_field': booking.courier_booking_date if booking.courier_booking_date else '',
                'courier_booking_date_val': booking.booking_date if booking.booking_date else '',
                'courier_dc_date': booking.dc_date if booking.dc_date else '',
                'courier_partner_name': booking.courier_partner_id.name if booking.courier_partner_id else '',
                'courier_booking_number': booking.courier_booking_number or '',
            })

        return preview_data

    def generate_courier_report(self):
        """Generate Courier Report Excel file"""
        domain = [('dispatch_type', '=', 'courier')]

        if self.courier_from_date:
            domain.append(('courier_booking_date', '>=', self.courier_from_date))

        if self.courier_to_date:
            domain.append(('courier_booking_date', '<=', self.courier_to_date))

        if self.courier_partner_filter_id:
            domain.append(('courier_partner_id', '=', self.courier_partner_filter_id.id))

        if self.courier_order_type:
            domain.append(('order_type', '=', self.courier_order_type))

        bookings = self.env['delivery.booking'].search(domain, order='courier_booking_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Courier Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('DC Number',              20),
            ('Sale Number',            25),
            ('Order Type',             15),
            ('Customer',               30),
            ('Customer PO Number',     22),
            ('PO Issue Date',          15),
            ('Courier Booking Date',   20),
            ('Delivery Booking Date',  20),
            ('DC Date',                15),
            ('Courier Name',           25),
            ('Courier Booking Number', 25),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        row = 1
        order_type_map = {
            'elcot': 'ELCOT Order',
            'others': 'Direct Order',
        }
        for booking in bookings:
            order_type_label = order_type_map.get(booking.order_type, booking.order_type or '')

            dc_number = ''
            customer = ''
            if booking.dc_id:
                dc_number = booking.dc_id.dc_number or ''
                customer = booking.dc_id.partner_id.name if booking.dc_id.partner_id else ''

            data = [
                dc_number,
                booking.sale_order_id.name if booking.sale_order_id else '',
                order_type_label,
                customer,
                booking.sale_order_id.client_order_ref or '' if booking.sale_order_id else '',
                booking.sale_order_id.po_issue_date if booking.sale_order_id and booking.sale_order_id.po_issue_date else '',
                booking.courier_booking_date if booking.courier_booking_date else '',
                booking.booking_date if booking.booking_date else '',
                # booking.dc_date.date() if booking.dc_date else '',
                booking.dc_date if booking.dc_date else '',
                booking.courier_partner_id.name if booking.courier_partner_id else '',
                booking.courier_booking_number or '',
            ]

            date_cols = {5, 6, 7, 8}
            for col, val in enumerate(data):
                fmt = date_format if col in date_cols else cell_format
                worksheet.write(row, col, val, fmt)

            row += 1

        workbook.close()
        output.seek(0)

        filename = f'courier_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }


    def _get_overview_preview_data(self):
        """Get preview data - one row per sale order"""

        enq_status_map = {
            'new': 'New', 'open': 'Open', 'reject': 'Closed', 'contacted': 'Converted To Lead',
        }
        lead_status_map = {
            'pending': 'Pending', 'confirm': 'Quoted', 'cancel': 'Canceled',
        }
        quot_status_map = {
            'draft': 'Draft', 'negotiate': 'Negotiation', 'waiting_approval': 'Waiting for Approval',
            'rejected': 'Rejected', 'approved': 'Approved', 'quote_sent': 'Quotation Sent',
            'resend_for_approval': 'Resend for Approval', 'cancel': 'Cancelled',
            'confirm': 'Order Confirmed', 'inactive': 'Inactive',
        }
        sale_status_map = {
            'draft': 'Workorder Draft', 'sent': 'Workorder Approval',
            'sale': 'Workorder Order', 'cancel': 'Cancelled',
        }
        dc_status_map = {
            'draft': 'Draft', 'waiting': 'Waiting Another Operation', 'confirmed': 'Waiting',
            'assigned': 'Inspection', 'outward': 'Outward', 'inward': 'Inward',
            'done': 'Done', 'cancel': 'Cancelled',
        }
        order_type_map = {
            'elcot': 'ELCOT Order', 'others': 'Direct Order',
        }

        enq_domain = []
        if self.from_date:
            enq_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            enq_domain.append(('create_date', '<=', self.to_date))
        if self.customer_id:
            enq_domain.append(('company_id', '=', self.customer_id.id))
        enquiries = self.env['customer.enq'].search(enq_domain, order='create_date asc')

        lead_domain = []
        if self.from_date:
            lead_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            lead_domain.append(('create_date', '<=', self.to_date))
        all_leads = self.env['order.enq'].search(lead_domain, order='create_date asc')

        quot_domain = []
        if self.from_date:
            quot_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            quot_domain.append(('create_date', '<=', self.to_date))
        all_quotations = self.env['quotation.management'].search(quot_domain, order='create_date asc')

        sale_domain = []
        if self.from_date:
            sale_domain.append(('date_order', '>=', self.from_date))
        if self.to_date:
            sale_domain.append(('date_order', '<=', self.to_date))
        all_sales = self.env['sale.order'].search(sale_domain, order='date_order asc')

        all_consignees = self.env['consignee.separation'].search([
            ('sale_order_id', 'in', all_sales.ids)
        ]) if all_sales else self.env['consignee.separation']

        all_dcs = self.env['stock.picking'].search([
            ('picking_type_code', '=', 'outgoing'),
            ('state', '!=', 'cancel'),
            ('sale_id', 'in', all_sales.ids)
        ], order='name asc') if all_sales else self.env['stock.picking']

        all_invoices = self.env['account.move'].search([
            ('move_type', '=', 'out_invoice'),
            ('work_order_id', 'in', all_sales.ids),
        ], order='name asc') if all_sales else self.env['account.move']

        all_tasks = self.env['project.task'].search([
            ('sale_order_id', 'in', all_sales.ids)
        ], order='name asc') if all_sales else self.env['project.task']

        # Lookup maps
        lead_by_enq = {l.enq_id.id: l for l in all_leads if l.enq_id}
        quot_by_enq = {q.enq_id.id: q for q in all_quotations if q.enq_id}

        # KEY FIX: one quotation → MANY sales (list)
        sales_by_quot = {}
        for s in all_sales:
            if s.quotation_id:
                sales_by_quot.setdefault(s.quotation_id.id, []).append(s)

        # sale_id → quotation (reverse lookup)
        quot_by_sale = {s.id: s.quotation_id for s in all_sales if s.quotation_id}

        consignees_by_sale = {}
        for c in all_consignees:
            consignees_by_sale.setdefault(c.sale_order_id.id, []).append(c)

        dcs_by_sale = {}
        for dc in all_dcs:
            dcs_by_sale.setdefault(dc.sale_id.id, []).append(dc)

        invoices_by_sale = {}
        for inv in all_invoices:
            if inv.work_order_id:
                invoices_by_sale.setdefault(inv.work_order_id.id, []).append(inv)

        tasks_by_sale = {}
        for task in all_tasks:
            if task.sale_order_id:
                tasks_by_sale.setdefault(task.sale_order_id.id, []).append(task)

        invoice_status_map = {
            'draft': 'Draft', 'posted': 'Posted', 'cancel': 'Cancelled',
        }

        preview_data = []
        covered_sales = set()

        # Pass 1: Start from enquiries, expand to ALL linked sales
        for enq in enquiries:
            lead = lead_by_enq.get(enq.id)
            quotation = quot_by_enq.get(enq.id)

            # Get ALL sales linked to this quotation
            sales = sales_by_quot.get(quotation.id, []) if quotation else []

            if sales:
                for sale in sales:
                    if self.salesperson_id and sale.user_id != self.salesperson_id:
                        continue
                    covered_sales.add(sale.id)
                    consignees = consignees_by_sale.get(sale.id, [])
                    dcs = dcs_by_sale.get(sale.id, [])
                    invoices = invoices_by_sale.get(sale.id, [])
                    tasks = tasks_by_sale.get(sale.id, [])
                    base = {
                        'enquiry_number': enq.name or '',
                        'enquiry_date': enq.create_date.date() if enq.create_date else False,
                        'enquiry_status': enq_status_map.get(enq.status, enq.status or ''),
                        'lead_number': lead.name if lead else '',
                        'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                        'lead_sales_person': sale.user_id.name if sale.user_id else '',
                        'lead_customer_name': self._get_customer_name(enq, sale),
                        'order_type': order_type_map.get(sale.order_type, sale.order_type or '') if sale else '',
                        'quotation_number': quotation.name if quotation else '',
                        'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                        'sale_order_number': sale.name or '',
                        'sale_order_status': sale_status_map.get(sale.state, sale.state or '') if sale else '',
                    }
                    items = list(sale.item_line_ids)
                    max_rows = max(len(items), len(dcs), len(invoices), len(tasks), 1)
                    for i in range(max_rows):
                        item = items[i] if i < len(items) else None
                        dc = dcs[i] if i < len(dcs) else None
                        invoice = invoices[i] if i < len(invoices) else None
                        task = tasks[i] if i < len(tasks) else None
                        if dc:
                            dc_number_v = getattr(dc, 'dc_number', None) or dc.name
                            dc_qty_v = float(sum(il.order_qty for il in dc.item_line_ids))
                            dc_status_v = dc_status_map.get(dc.state, dc.state or '')
                            dc_total_v = dc.amount_total or 0
                            consignee_v = dc.consignee_separation_id.name or '' if dc.consignee_separation_id else ''
                        else:
                            dc_number_v = ''
                            dc_qty_v = 0.0
                            dc_status_v = ''
                            dc_total_v = 0
                            consignee_v = ''
                        preview_data.append({
                            **base,
                            'overview_item_names': item.item_name or '' if item else '',
                            'overview_item_descriptions': item.description or '' if item else '',
                            'overview_item_quantities': str(item.order_qty) if item else '',
                            'consignee_numbers': consignee_v,
                            'dc_numbers': dc_number_v,
                            'overview_dc_quantity': dc_qty_v,
                            'dc_statuses': dc_status_v,
                            'dc_total': dc_total_v,
                            'invoice_numbers': invoice.name or '' if invoice else '',
                            'invoice_statuses': invoice_status_map.get(invoice.state, invoice.state or '') if invoice else '',
                            'invoice_total': (invoice.item_total_amount or invoice.amount_total) if invoice else 0,
                            'installation_numbers': task.name or '' if task else '',
                            'installation_statuses': task.stage_id.name if task and task.stage_id else '',
                        })
            else:
                # Enquiry/lead/quotation exists but no sale yet
                if self.salesperson_id:
                    continue
                preview_data.append({
                    'enquiry_number': enq.name or '',
                    'enquiry_date': enq.create_date.date() if enq.create_date else False,
                    'enquiry_status': enq_status_map.get(enq.status, enq.status or ''),
                    'lead_number': lead.name if lead else '',
                    'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                    'lead_sales_person': '',
                    'lead_customer_name': self._get_customer_name(enq, None),
                    'order_type': '',
                    'quotation_number': quotation.name if quotation else '',
                    'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                    'sale_order_number': '',
                    'sale_order_status': '',
                    'overview_item_names': '',
                    'overview_item_descriptions': '',
                    'overview_item_quantities': '',
                    'consignee_numbers': '',
                    'dc_numbers': '',
                    'overview_dc_quantity': 0.0,
                    'dc_statuses': '',
                    'dc_total': 0,
                    'invoice_numbers': '',
                    'invoice_statuses': '',
                    'invoice_total': 0,
                    'installation_numbers': '',
                    'installation_statuses': '',
                })

        # Pass 2: Sales NOT yet covered - trace back to quotation/enquiry
        for sale in all_sales:
            if sale.id in covered_sales:
                continue
            if self.salesperson_id and sale.user_id != self.salesperson_id:
                continue

            covered_sales.add(sale.id)

            # Trace back: sale → quotation → enquiry → lead
            quotation = quot_by_sale.get(sale.id)
            enq = None
            lead = None
            if quotation:
                enq_id = quotation.enq_id.id if quotation.enq_id else None
                if enq_id:
                    enq = quotation.enq_id
                    lead = lead_by_enq.get(enq_id)

            consignees = consignees_by_sale.get(sale.id, [])
            dcs = dcs_by_sale.get(sale.id, [])
            invoices = invoices_by_sale.get(sale.id, [])
            tasks = tasks_by_sale.get(sale.id, [])

            base = {
                'enquiry_number': enq.name if enq else '',
                'enquiry_date': enq.create_date.date() if enq and enq.create_date else False,
                'enquiry_status': enq_status_map.get(enq.status, enq.status or '') if enq else '',
                'lead_number': lead.name if lead else '',
                'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                'lead_sales_person': sale.user_id.name if sale.user_id else '',
                'lead_customer_name': self._get_customer_name(enq, sale),
                'order_type': order_type_map.get(sale.order_type, sale.order_type or '') if sale else '',
                'quotation_number': quotation.name if quotation else '',
                'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                'sale_order_number': sale.name or '',
                'sale_order_status': sale_status_map.get(sale.state, sale.state or '') if sale else '',
            }
            items = list(sale.item_line_ids)
            max_rows = max(len(items), len(dcs), len(invoices), len(tasks), 1)
            for i in range(max_rows):
                item = items[i] if i < len(items) else None
                dc = dcs[i] if i < len(dcs) else None
                invoice = invoices[i] if i < len(invoices) else None
                task = tasks[i] if i < len(tasks) else None
                if dc:
                    dc_number_v = getattr(dc, 'dc_number', None) or dc.name
                    dc_qty_v = float(sum(il.order_qty for il in dc.item_line_ids))
                    dc_status_v = dc_status_map.get(dc.state, dc.state or '')
                    dc_total_v = dc.amount_total or 0
                    consignee_v = dc.consignee_separation_id.name or '' if dc.consignee_separation_id else ''
                else:
                    dc_number_v = ''
                    dc_qty_v = 0.0
                    dc_status_v = ''
                    dc_total_v = 0
                    consignee_v = ''
                preview_data.append({
                    **base,
                    'overview_item_names': item.item_name or '' if item else '',
                    'overview_item_descriptions': item.description or '' if item else '',
                    'overview_item_quantities': str(item.order_qty) if item else '',
                    'consignee_numbers': consignee_v,
                    'dc_numbers': dc_number_v,
                    'overview_dc_quantity': dc_qty_v,
                    'dc_statuses': dc_status_v,
                    'dc_total': dc_total_v,
                    'invoice_numbers': invoice.name or '' if invoice else '',
                    'invoice_statuses': invoice_status_map.get(invoice.state, invoice.state or '') if invoice else '',
                    'invoice_total': (invoice.item_total_amount or invoice.amount_total) if invoice else 0,
                    'installation_numbers': task.name or '' if task else '',
                    'installation_statuses': task.stage_id.name if task and task.stage_id else '',
                })

        return preview_data


    def generate_overview_report(self):
        """Generate Overview Report - one row per sale order, tracing back to enquiry"""

        enq_status_map = {
            'new': 'New', 'open': 'Open', 'reject': 'Closed', 'contacted': 'Converted To Lead',
        }
        lead_status_map = {
            'pending': 'Pending', 'confirm': 'Quoted', 'cancel': 'Canceled',
        }
        quot_status_map = {
            'draft': 'Draft', 'negotiate': 'Negotiation', 'waiting_approval': 'Waiting for Approval',
            'rejected': 'Rejected', 'approved': 'Approved', 'quote_sent': 'Quotation Sent',
            'resend_for_approval': 'Resend for Approval', 'cancel': 'Cancelled',
            'confirm': 'Order Confirmed', 'inactive': 'Inactive',
        }
        sale_status_map = {
            'draft': 'Workorder Draft', 'sent': 'Workorder Approval',
            'sale': 'Workorder Order', 'cancel': 'Cancelled',
        }
        dc_status_map = {
            'draft': 'Draft', 'waiting': 'Waiting Another Operation', 'confirmed': 'Waiting',
            'assigned': 'Inspection', 'outward': 'Outward', 'inward': 'Inward',
            'done': 'Done', 'cancel': 'Cancelled',
        }
        order_type_map = {
            'elcot': 'ELCOT Order', 'others': 'Direct Order',
        }

        # 1. All Enquiries
        enq_domain = []
        if self.from_date:
            enq_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            enq_domain.append(('create_date', '<=', self.to_date))
        if self.customer_id:
            enq_domain.append(('company_id', '=', self.customer_id.id))
        enquiries = self.env['customer.enq'].search(enq_domain, order='create_date asc')

        # 2. All Leads
        lead_domain = []
        if self.from_date:
            lead_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            lead_domain.append(('create_date', '<=', self.to_date))
        all_leads = self.env['order.enq'].search(lead_domain, order='create_date asc')

        # 3. All Quotations
        quot_domain = []
        if self.from_date:
            quot_domain.append(('create_date', '>=', self.from_date))
        if self.to_date:
            quot_domain.append(('create_date', '<=', self.to_date))
        all_quotations = self.env['quotation.management'].search(quot_domain, order='create_date asc')

        # 4. All Sale Orders
        sale_domain = []
        if self.from_date:
            sale_domain.append(('date_order', '>=', self.from_date))
        if self.to_date:
            sale_domain.append(('date_order', '<=', self.to_date))
        all_sales = self.env['sale.order'].search(sale_domain, order='date_order asc')

        # 5. All Consignee Separations
        all_consignees = self.env['consignee.separation'].search([
            ('sale_order_id', 'in', all_sales.ids)
        ]) if all_sales else self.env['consignee.separation']

        # 6. All DCs
        all_dcs = self.env['stock.picking'].search([
            ('picking_type_code', '=', 'outgoing'),
            ('state', '!=', 'cancel'),
            ('sale_id', 'in', all_sales.ids)
        ], order='name asc') if all_sales else self.env['stock.picking']

        # 7. All Invoices
        all_invoices = self.env['account.move'].search([
            ('move_type', '=', 'out_invoice'),
            ('work_order_id', 'in', all_sales.ids),
        ], order='name asc') if all_sales else self.env['account.move']

        # 8. All Installation Tasks
        all_tasks = self.env['project.task'].search([
            ('sale_order_id', 'in', all_sales.ids)
        ], order='name asc') if all_sales else self.env['project.task']

        # ── Lookup maps ───────────────────────────────────────────────
        lead_by_enq = {l.enq_id.id: l for l in all_leads if l.enq_id}
        quot_by_enq = {q.enq_id.id: q for q in all_quotations if q.enq_id}

        # One quotation → MANY sales
        sales_by_quot = {}
        for s in all_sales:
            if s.quotation_id:
                sales_by_quot.setdefault(s.quotation_id.id, []).append(s)

        # Reverse: sale → quotation
        quot_by_sale = {s.id: s.quotation_id for s in all_sales if s.quotation_id}

        consignees_by_sale = {}
        for c in all_consignees:
            consignees_by_sale.setdefault(c.sale_order_id.id, []).append(c)

        dcs_by_sale = {}
        for dc in all_dcs:
            dcs_by_sale.setdefault(dc.sale_id.id, []).append(dc)

        invoices_by_sale = {}
        for inv in all_invoices:
            if inv.work_order_id:
                invoices_by_sale.setdefault(inv.work_order_id.id, []).append(inv)

        tasks_by_sale = {}
        for task in all_tasks:
            if task.sale_order_id:
                tasks_by_sale.setdefault(task.sale_order_id.id, []).append(task)

        invoice_status_map = {
            'draft': 'Draft', 'posted': 'Posted', 'cancel': 'Cancelled',
        }

        rows = []
        covered_sales = set()

        # ── Pass 1: From enquiries → expand to ALL linked sales ───────
        for enq in enquiries:
            lead = lead_by_enq.get(enq.id)
            quotation = quot_by_enq.get(enq.id)
            sales = sales_by_quot.get(quotation.id, []) if quotation else []

            if sales:
                for sale in sales:
                    if self.salesperson_id and sale.user_id != self.salesperson_id:
                        continue
                    covered_sales.add(sale.id)
                    consignees = consignees_by_sale.get(sale.id, [])
                    dcs = dcs_by_sale.get(sale.id, [])
                    invoices = invoices_by_sale.get(sale.id, [])
                    tasks = tasks_by_sale.get(sale.id, [])
                    base = {
                        'enquiry_number': enq.name or '',
                        'enquiry_date': enq.create_date.date() if enq.create_date else '',
                        'enquiry_status': enq_status_map.get(enq.status, enq.status or ''),
                        'lead_number': lead.name if lead else '',
                        'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                        'salesperson': sale.user_id.name if sale.user_id else '',
                        'customer': self._get_customer_name(enq, sale),
                        'order_type': order_type_map.get(sale.order_type, sale.order_type or '') if sale else '',
                        'quotation_number': quotation.name if quotation else '',
                        'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                        'quotation_date': quotation.quotation_date if quotation else '',
                        'sale_order_number': sale.name or '',
                        'customer_po_number': sale.client_order_ref or '',
                        'sale_order_status': sale_status_map.get(sale.state, sale.state or '') if sale else '',
                        'sale_order_date': (sale.date_order.date() if isinstance(sale.date_order, datetime) else sale.date_order) if sale.date_order else '',
                    }
                    items = list(sale.item_line_ids)
                    max_rows = max(len(items), len(dcs), len(invoices), len(tasks), 1)
                    for i in range(max_rows):
                        item = items[i] if i < len(items) else None
                        dc = dcs[i] if i < len(dcs) else None
                        invoice = invoices[i] if i < len(invoices) else None
                        task = tasks[i] if i < len(tasks) else None
                        if dc:
                            dc_number_v = getattr(dc, 'dc_number', None) or dc.name
                            dc_qty_v = sum(il.order_qty for il in dc.item_line_ids)
                            dc_status_v = dc_status_map.get(dc.state, dc.state or '')
                            delivery_status_v = 'Delivered' if dc.state == 'done' else 'Undelivered'
                            _d = dc.scheduled_date or dc.date_done
                            dc_date_v = _d.strftime('%d/%m/%Y') if _d else ''
                            dc_total_v = dc.amount_total or 0
                            consignee_v = dc.consignee_separation_id.name or '' if dc.consignee_separation_id else ''
                            consignee_value_v = dc.consignee_separation_id.amount_total or 0 if dc.consignee_separation_id else 0
                        else:
                            dc_number_v = ''
                            dc_qty_v = 0
                            dc_status_v = ''
                            delivery_status_v = ''
                            dc_date_v = ''
                            dc_total_v = 0
                            consignee_v = ''
                            consignee_value_v = 0
                        rows.append({
                            **base,
                            'item_names': item.item_name or '' if item else '',
                            'item_descriptions': item.description or '' if item else '',
                            'item_unit_price': item.unit_price if item else 0,
                            'item_quantities': str(item.order_qty) if item else '',
                            'item_subtotal': item.price_subtotal if item else 0,
                            'consignee_numbers': consignee_v,
                            'consignee_value': consignee_value_v,
                            'dc_numbers': dc_number_v,
                            'dc_quantity': dc_qty_v,
                            'dc_statuses': dc_status_v,
                            'delivery_status': delivery_status_v,
                            'dc_dates': dc_date_v,
                            'dc_total': dc_total_v,
                            'invoice_numbers': invoice.name or '' if invoice else '',
                            'invoice_statuses': invoice_status_map.get(invoice.state, invoice.state or '') if invoice else '',
                            'invoice_total': (invoice.item_total_amount or invoice.amount_total) if invoice else 0,
                            'installation_numbers': task.name or '' if task else '',
                            'installation_statuses': task.stage_id.name if task and task.stage_id else '',
                            'installation_amount': sum(task.item_delivery_ids.mapped('price_total')) if task else 0,
                        })
            else:
                # Enquiry exists but no sale yet - still show the row
                if self.salesperson_id:
                    continue
                rows.append({
                    'enquiry_number': enq.name or '',
                    'enquiry_date': enq.create_date.date() if enq.create_date else '',
                    'enquiry_status': enq_status_map.get(enq.status, enq.status or ''),
                    'lead_number': lead.name if lead else '',
                    'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                    'salesperson': '',
                    'customer': self._get_customer_name(enq, None),
                    'order_type': '',
                    'quotation_number': quotation.name if quotation else '',
                    'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                    'quotation_date': quotation.quotation_date if quotation else '',
                    'sale_order_number': '',
                    'customer_po_number': '',
                    'sale_order_status': '',
                    'sale_order_date': '',
                    'item_names': '',
                    'item_descriptions': '',
                    'item_unit_price': 0,
                    'item_quantities': '',
                    'item_subtotal': 0,
                    'consignee_numbers': '',
                    'consignee_value': 0,
                    'dc_numbers': '',
                    'dc_quantity': 0,
                    'dc_statuses': '',
                    'delivery_status': '',
                    'dc_dates': '',
                    'dc_total': 0,
                    'invoice_numbers': '',
                    'invoice_statuses': '',
                    'invoice_total': 0,
                    'installation_numbers': '',
                    'installation_statuses': '',
                    'installation_amount': 0,
                })

        # ── Pass 2: Remaining sales - trace back to quotation/enquiry ─
        for sale in all_sales:
            if sale.id in covered_sales:
                continue
            if self.salesperson_id and sale.user_id != self.salesperson_id:
                continue

            covered_sales.add(sale.id)
            quotation = quot_by_sale.get(sale.id)
            enq = None
            lead = None
            if quotation and quotation.enq_id:
                enq = quotation.enq_id
                lead = lead_by_enq.get(enq.id)

            consignees = consignees_by_sale.get(sale.id, [])
            dcs = dcs_by_sale.get(sale.id, [])
            invoices = invoices_by_sale.get(sale.id, [])
            tasks = tasks_by_sale.get(sale.id, [])

            base = {
                'enquiry_number': enq.name if enq else '',
                'enquiry_date': enq.create_date.date() if enq and enq.create_date else '',
                'enquiry_status': enq_status_map.get(enq.status, enq.status or '') if enq else '',
                'lead_number': lead.name if lead else '',
                'lead_status': lead_status_map.get(lead.state, lead.state or '') if lead else '',
                'salesperson': sale.user_id.name if sale.user_id else '',
                'customer': self._get_customer_name(enq, sale),
                'order_type': order_type_map.get(sale.order_type, sale.order_type or '') if sale else '',
                'quotation_number': quotation.name if quotation else '',
                'quotation_status': quot_status_map.get(quotation.state, quotation.state or '') if quotation else '',
                'quotation_date': quotation.quotation_date if quotation else '',
                'sale_order_number': sale.name or '',
                'customer_po_number': sale.client_order_ref or '',
                'sale_order_status': sale_status_map.get(sale.state, sale.state or '') if sale else '',
                'sale_order_date': (sale.date_order.date() if isinstance(sale.date_order, datetime) else sale.date_order) if sale.date_order else '',
            }
            items = list(sale.item_line_ids)
            max_rows = max(len(items), len(dcs), len(invoices), len(tasks), 1)
            for i in range(max_rows):
                item = items[i] if i < len(items) else None
                dc = dcs[i] if i < len(dcs) else None
                invoice = invoices[i] if i < len(invoices) else None
                task = tasks[i] if i < len(tasks) else None
                if dc:
                    dc_number_v = getattr(dc, 'dc_number', None) or dc.name
                    dc_qty_v = sum(il.order_qty for il in dc.item_line_ids)
                    dc_status_v = dc_status_map.get(dc.state, dc.state or '')
                    delivery_status_v = 'Delivered' if dc.state == 'done' else 'Undelivered'
                    _d = dc.scheduled_date or dc.date_done
                    dc_date_v = _d.strftime('%d/%m/%Y') if _d else ''
                    dc_total_v = dc.amount_total or 0
                    consignee_v = dc.consignee_separation_id.name or '' if dc.consignee_separation_id else ''
                    consignee_value_v = dc.consignee_separation_id.amount_total or 0 if dc.consignee_separation_id else 0
                else:
                    dc_number_v = ''
                    dc_qty_v = 0
                    dc_status_v = ''
                    delivery_status_v = ''
                    dc_date_v = ''
                    dc_total_v = 0
                    consignee_v = ''
                    consignee_value_v = 0
                rows.append({
                    **base,
                    'item_names': item.item_name or '' if item else '',
                    'item_descriptions': item.description or '' if item else '',
                    'item_unit_price': item.unit_price if item else 0,
                    'item_quantities': str(item.order_qty) if item else '',
                    'item_subtotal': item.price_subtotal if item else 0,
                    'consignee_numbers': consignee_v,
                    'consignee_value': consignee_value_v,
                    'dc_numbers': dc_number_v,
                    'dc_quantity': dc_qty_v,
                    'dc_statuses': dc_status_v,
                    'delivery_status': delivery_status_v,
                    'dc_dates': dc_date_v,
                    'dc_total': dc_total_v,
                    'invoice_numbers': invoice.name or '' if invoice else '',
                    'invoice_statuses': invoice_status_map.get(invoice.state, invoice.state or '') if invoice else '',
                    'invoice_total': (invoice.item_total_amount or invoice.amount_total) if invoice else 0,
                    'installation_numbers': task.name or '' if task else '',
                    'installation_statuses': task.stage_id.name if task and task.stage_id else '',
                    'installation_amount': sum(task.item_delivery_ids.mapped('price_total')) if task else 0,
                })

        # ── Write Excel ───────────────────────────────────────────────
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Overview Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter'
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter'
        })

        columns = [
            ('Enquiry Number', 20),
            ('Enquiry Date', 15),
            ('Enquiry Status', 20),
            ('Lead Number', 20),
            ('Lead Status', 18),
            ('Salesperson', 20),
            ('Customer', 25),
            ('Order Type', 18),
            ('Quotation Number', 20),
            ('Quotation Status', 22),
            ('Quotation Date', 15),
            ('Sale Order Number', 20),
            ('Customer PO Number', 22),
            ('Sale Order Status', 22),
            ('Sale Order Date', 15),
            ('Item Name', 30),
            ('Item Description', 35),
            ('Unit Price', 18),
            ('Item Quantity', 18),
            ('Item Tax Excl. Amount', 22),
            ('Consignee Numbers', 25),
            ('Consignee Value', 22),
            ('DC Number', 25),
            ('DC Quantity', 18),
            ('DC Status', 25),
            ('Delivered/Undelivered', 22),
            ('DC Date', 20),
            ('DC Total Amount', 20),
            ('Invoice Number', 25),
            ('Invoice Status', 18),
            ('Invoice Total Amount', 22),
            ('Installation Number', 25),
            ('Installation Status', 22),
            ('Installation Amount', 22),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        amount_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter'
        })

        date_cols = {1, 10, 14}
        amount_cols = {17, 19, 21, 23, 27, 30, 33}
        for r, row in enumerate(rows, start=1):
            values = [
                row['enquiry_number'],       # 0
                row['enquiry_date'],          # 1  date
                row['enquiry_status'],        # 2
                row['lead_number'],           # 3
                row['lead_status'],           # 4
                row['salesperson'],           # 5
                row['customer'],              # 6
                row['order_type'],            # 7
                row['quotation_number'],      # 8
                row['quotation_status'],      # 9
                row['quotation_date'],        # 10 date
                row['sale_order_number'],     # 11
                row['customer_po_number'],    # 12
                row['sale_order_status'],     # 13
                row['sale_order_date'],       # 14 date
                row['item_names'],            # 14
                row['item_descriptions'],     # 15
                row['item_unit_price'],       # 16 amount
                row['item_quantities'],       # 17
                row['item_subtotal'],         # 18 amount
                row['consignee_numbers'],     # 19
                row['consignee_value'],       # 20 amount
                row['dc_numbers'],            # 21
                row['dc_quantity'],           # 21 amount
                row['dc_statuses'],           # 22
                row['delivery_status'],       # 23
                row['dc_dates'],              # 24
                row['dc_total'],              # 25 amount
                row['invoice_numbers'],       # 26
                row['invoice_statuses'],      # 27
                row['invoice_total'],         # 28 amount
                row['installation_numbers'],  # 29
                row['installation_statuses'], # 30
                row['installation_amount'],   # 31 amount
            ]
            for col, val in enumerate(values):
                if col in date_cols and val:
                    worksheet.write(r, col, val, date_format)
                elif col in amount_cols:
                    worksheet.write(r, col, val or 0, amount_format)
                else:
                    worksheet.write(r, col, val or '', cell_format)

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.from_date and self.to_date:
            filter_parts.append(f'{self.from_date}_{self.to_date}')
        if self.customer_id:
            filter_parts.append(f'customer_{self.customer_id.name.replace(" ", "_")}')

        filename = f'overview_report_{"_".join(filter_parts) if filter_parts else "all"}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }
    
    def _get_delivery_booking_summary_preview_data(self):
        """Get preview data for Delivery Booking Summary Report"""
        domain = []

        if self.delivery_booking_from_date:
            domain.append(('booking_date', '>=', self.delivery_booking_from_date))

        if self.delivery_booking_to_date:
            domain.append(('booking_date', '<=', self.delivery_booking_to_date))

        if self.delivery_booking_customer_id:
            domain.append(('customer_id', '=', self.delivery_booking_customer_id.id))

        if self.delivery_booking_dispatch_type:
            domain.append(('dispatch_type', '=', self.delivery_booking_dispatch_type))

        if self.delivery_booking_status:
            domain.append(('state', '=', self.delivery_booking_status))

        bookings = self.env['delivery.booking'].search(domain, order='booking_date asc')

        dispatch_type_map = {
            'direct': 'Direct',
            'courier': 'Courier',
        }
        state_map = {
            'draft': 'Draft',
            'confirmed': 'Confirmed',
            'dispatched': 'Dispatched',
            'delivered': 'Delivered',
            'cancel': 'Cancelled',
        }

        def _format_addr(p):
            if not p:
                return ''
            parts = [
                p.street or '',
                p.street2 or '',
                p.city or '',
                p.state_id.name if p.state_id else '',
                p.zip or '',
                p.country_id.name if p.country_id else '',
            ]
            return ', '.join(x for x in parts if x)

        preview_data = []

        for booking in bookings:
            dc_number = ''
            dispatch_location = ''
            shipping_partner = False
            if booking.dc_id:
                dc_number = booking.dc_id.dc_number or booking.dc_id.name or ''
                warehouse = booking.dc_id.picking_type_id.warehouse_id
                dispatch_location = warehouse.name if warehouse else ''
                shipping_partner = booking.dc_id.partner_id

            billing_partner = booking.customer_id.commercial_partner_id if booking.customer_id else False
            contact_person = shipping_partner.name if shipping_partner else ''
            contact_no = (shipping_partner.phone or shipping_partner.mobile or '') if shipping_partner else ''
            # Find invoices linked to this DC directly (via selected_dc_id)
            # or via the sale order as a fallback
            dc_invoices = self.env['account.move'].search([
                ('selected_dc_id', '=', booking.dc_id.id),
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
            ]) if booking.dc_id else self.env['account.move']
            if not dc_invoices and booking.sale_order_id:
                dc_invoices = booking.sale_order_id.invoice_ids.filtered(
                    lambda i: i.state == 'posted' and i.move_type == 'out_invoice'
                )
            invoice_nos = ', '.join(inv.name for inv in dc_invoices)

            # Build item value map and unit price map from DC item lines
            item_value_map = {}
            unit_price_map = {}
            if booking.dc_id:
                for dc_item in booking.dc_id.item_line_ids:
                    key = dc_item.item_details_id.id if dc_item.item_details_id else 0
                    item_value_map[key] = item_value_map.get(key, 0.0) + (dc_item.price_subtotal or 0.0)
                    if key not in unit_price_map:
                        unit_price_map[key] = dc_item.price_unit or 0.0

            base = {
                'db_booking_number': booking.name or '',
                'db_dc_number': dc_number,
                'db_sale_number': booking.sale_order_id.name if booking.sale_order_id else '',
                'db_customer_po_number': booking.sale_order_id.client_order_ref or '' if booking.sale_order_id else '',
                'db_po_issue_date': booking.sale_order_id.po_issue_date if booking.sale_order_id and booking.sale_order_id.po_issue_date else False,
                'db_dispatch_location': dispatch_location,
                'db_customer': billing_partner.name if billing_partner else '',
                'db_billing_address': _format_addr(billing_partner),
                'db_shipping_address': _format_addr(shipping_partner),
                'db_contact_person': contact_person,
                'db_contact_no': contact_no,
                'db_booking_date': booking.booking_date if booking.booking_date else False,
                'db_dispatch_type': dispatch_type_map.get(booking.dispatch_type, booking.dispatch_type or ''),
                'db_surface': {'land': 'Land', 'air': 'Air'}.get(booking.surface, '') if booking.surface else '',
                'db_driver_name': booking.driver_name or '',
                'db_vehicle_number': booking.vehicle_number or '',
                'db_status': state_map.get(booking.state, booking.state or ''),
                'db_district': shipping_partner.city or '' if shipping_partner else '',
                'db_invoice_no': invoice_nos,
            }

            if booking.item_line_ids:
                for item_line in booking.item_line_ids:
                    key = item_line.item_details_id.id if item_line.item_details_id else 0
                    row_data = dict(base)
                    row_data['db_item_details'] = item_line.item_name or ''
                    row_data['db_item_description'] = item_line.description or ''
                    row_data['db_item_qty'] = str(int(item_line.quantity)) if item_line.quantity else ''
                    row_data['db_unit_price'] = unit_price_map.get(key, 0.0)
                    row_data['db_dc_value'] = item_value_map.get(key, 0.0)
                    preview_data.append(row_data)
            else:
                base['db_item_details'] = ''
                base['db_item_description'] = ''
                base['db_item_qty'] = ''
                base['db_unit_price'] = 0.0
                base['db_dc_value'] = 0.0
                preview_data.append(base)

        return preview_data


    def generate_delivery_booking_summary_report(self):
        """Generate Delivery Booking Summary Report Excel file"""
        domain = []

        if self.delivery_booking_from_date:
            domain.append(('booking_date', '>=', self.delivery_booking_from_date))

        if self.delivery_booking_to_date:
            domain.append(('booking_date', '<=', self.delivery_booking_to_date))

        if self.delivery_booking_customer_id:
            domain.append(('customer_id', '=', self.delivery_booking_customer_id.id))

        if self.delivery_booking_dispatch_type:
            domain.append(('dispatch_type', '=', self.delivery_booking_dispatch_type))

        if self.delivery_booking_status:
            domain.append(('state', '=', self.delivery_booking_status))

        bookings = self.env['delivery.booking'].search(domain, order='booking_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Delivery Booking Summary')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Delivery Booking Number', 25),   # 0
            ('DC Number',               22),   # 1
            ('Sale Number',             25),   # 2
            ('Customer PO Number',      25),   # 3
            ('PO Issue Date',           15),   # 4  date
            ('Dispatch Location',       22),   # 5
            ('Customer Name',           30),   # 6
            ('Billing Address',         40),   # 7
            ('Shipping Address',        40),   # 8
            ('Contact Person',          22),   # 9
            ('Contact No',              18),   # 10
            ('Delivery Booking Date',   22),   # 11 date
            ('Dispatch Type',           18),   # 12
            ('Surface',                 15),   # 13
            ('Driver Name',             22),   # 14
            ('Vehicle Number',          18),   # 15
            ('Status',                  18),   # 16
            ('District',                20),   # 17
            ('Item Name',               35),   # 18
            ('Item Description',        35),   # 19
            ('Item Qty',                15),   # 20
            ('Unit Price',              18),   # 21
            ('DC Value',                18),   # 22
            ('Invoice No',              22),   # 23
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        dispatch_type_map = {
            'direct': 'Direct',
            'courier': 'Courier',
        }
        state_map = {
            'draft': 'Draft',
            'confirmed': 'Confirmed',
            'dispatched': 'Dispatched',
            'delivered': 'Delivered',
            'cancel': 'Cancelled',
        }

        def _format_addr(p):
            if not p:
                return ''
            parts = [
                p.street or '',
                p.street2 or '',
                p.city or '',
                p.state_id.name if p.state_id else '',
                p.zip or '',
                p.country_id.name if p.country_id else '',
            ]
            return ', '.join(x for x in parts if x)

        row = 1
        for booking in bookings:
            dc_number = ''
            dispatch_location = ''
            shipping_partner = False
            if booking.dc_id:
                dc_number = booking.dc_id.dc_number or booking.dc_id.name or ''
                warehouse = booking.dc_id.picking_type_id.warehouse_id
                dispatch_location = warehouse.name if warehouse else ''
                shipping_partner = booking.dc_id.partner_id

            billing_partner = booking.customer_id.commercial_partner_id if booking.customer_id else False
            contact_person = shipping_partner.name if shipping_partner else ''
            contact_no = (shipping_partner.phone or shipping_partner.mobile or '') if shipping_partner else ''
            # Find invoices linked to this DC directly (via selected_dc_id)
            # or via the sale order as a fallback
            dc_invoices = self.env['account.move'].search([
                ('selected_dc_id', '=', booking.dc_id.id),
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
            ]) if booking.dc_id else self.env['account.move']
            if not dc_invoices and booking.sale_order_id:
                dc_invoices = booking.sale_order_id.invoice_ids.filtered(
                    lambda i: i.state == 'posted' and i.move_type == 'out_invoice'
                )
            invoice_nos = ', '.join(inv.name for inv in dc_invoices)

            # Build item value map and unit price map from DC item lines
            item_value_map = {}
            unit_price_map = {}
            if booking.dc_id:
                for dc_item in booking.dc_id.item_line_ids:
                    key = dc_item.item_details_id.id if dc_item.item_details_id else 0
                    item_value_map[key] = item_value_map.get(key, 0.0) + (dc_item.price_subtotal or 0.0)
                    if key not in unit_price_map:
                        unit_price_map[key] = dc_item.price_unit or 0.0

            district = shipping_partner.city or '' if shipping_partner else ''

            common = [
                booking.name or '',
                dc_number,
                booking.sale_order_id.name if booking.sale_order_id else '',
                booking.sale_order_id.client_order_ref or '' if booking.sale_order_id else '',
                booking.sale_order_id.po_issue_date if booking.sale_order_id and booking.sale_order_id.po_issue_date else '',
                dispatch_location,
                billing_partner.name if billing_partner else '',
                _format_addr(billing_partner),
                _format_addr(shipping_partner),
                contact_person,
                contact_no,
                booking.booking_date if booking.booking_date else '',
                dispatch_type_map.get(booking.dispatch_type, booking.dispatch_type or ''),
                {'land': 'Land', 'air': 'Air'}.get(booking.surface, '') if booking.surface else '',
                booking.driver_name or '',
                booking.vehicle_number or '',
                state_map.get(booking.state, booking.state or ''),
                district,
            ]

            item_lines = booking.item_line_ids
            if not item_lines:
                data = common + ['', '', '', 0.0, 0.0, invoice_nos]
                for col, val in enumerate(data):
                    if col in {4, 11}:
                        fmt = date_format
                    elif col in {21, 22}:
                        fmt = number_format
                    else:
                        fmt = cell_format
                    worksheet.write(row, col, val, fmt)
                row += 1
            else:
                for item_line in item_lines:
                    key = item_line.item_details_id.id if item_line.item_details_id else 0
                    item_val = item_value_map.get(key, 0.0)
                    unit_price = unit_price_map.get(key, 0.0)
                    data = common + [
                        item_line.item_name or '',
                        item_line.description or '',
                        str(int(item_line.quantity)) if item_line.quantity else '',
                        unit_price,
                        item_val,
                        invoice_nos,
                    ]
                    for col, val in enumerate(data):
                        if col in {4, 11}:
                            fmt = date_format
                        elif col in {21, 22}:
                            fmt = number_format
                        else:
                            fmt = cell_format
                        worksheet.write(row, col, val, fmt)
                    row += 1

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.delivery_booking_from_date and self.delivery_booking_to_date:
            filter_parts.append(f'{self.delivery_booking_from_date}_{self.delivery_booking_to_date}')
        if self.delivery_booking_customer_id:
            filter_parts.append(f'customer_{self.delivery_booking_customer_id.name.replace(" ", "_")}')

        filename = (
            f'delivery_booking_summary_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }


    def _get_delivered_undelivered_preview_data(self):
        """Return preview rows for the Delivered/Undelivered Summary report."""
        domain = []
        if self.du_from_date:
            domain.append(('date_order', '>=', self.du_from_date))
        if self.du_to_date:
            domain.append(('date_order', '<=', self.du_to_date))
        if self.du_customer_id:
            domain.append(('partner_id', '=', self.du_customer_id.id))
        if self.du_workorder_ids:
            domain.append(('id', 'in', self.du_workorder_ids.ids))

        order_type_map = {'elcot': 'ELCOT Order', 'others': 'Direct Order'}
        preview_data = []

        for sale in self.env['sale.order'].search(domain, order='date_order asc'):
            done_dcs = self.env['stock.picking'].search([
                '|',
                ('sale_id', '=', sale.id),
                ('sale_order_id', '=', sale.id),
                ('picking_type_code', '=', 'outgoing'),
                ('state', '=', 'done'),
            ])
            inst_tasks = self.env['project.task'].search([('sale_order_id', '=', sale.id)])
            order_type = order_type_map.get(sale.order_type, sale.order_type or '')
            salesperson = sale.user_id.name if sale.user_id else ''
            posted_invoices = self.env['account.move'].search([
                ('work_order_id', '=', sale.id),
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
            ], order='invoice_date desc', limit=1)
            if posted_invoices:
                stc_val = posted_invoices.submitted_to_customer
                submitted_to_customer_str = 'Yes' if stc_val == 'yes' else ('No' if stc_val == 'no' else '')
                submitted_date = posted_invoices.submitted_date or False
            else:
                submitted_to_customer_str = ''
                submitted_date = False

            for item_line in sale.item_line_ids:
                workorder_qty = item_line.order_qty or 0.0
                unit_price = (
                    item_line.price_subtotal / item_line.order_qty
                    if item_line.order_qty else 0.0
                )
                total_amount = item_line.price_subtotal or 0.0
                tax_incl_total = item_line.price_total or 0.0
                tax_rate = (
                    (item_line.price_total / item_line.price_subtotal)
                    if item_line.price_subtotal else 1.0
                )
                delivered_qty = 0.0
                for dc in done_dcs:
                    for dc_item in dc.item_line_ids:
                        if (dc_item.item_details_id and item_line.item_details_id and
                                dc_item.item_details_id.id == item_line.item_details_id.id):
                            delivered_qty += dc_item.order_qty or 0.0
                delivered_amount = delivered_qty * unit_price
                delivered_amount_incl = delivered_amount * tax_rate
                undelivered_qty = max(0.0, workorder_qty - delivered_qty)
                undelivered_amount = undelivered_qty * unit_price
                undelivered_amount_incl = undelivered_qty * unit_price * tax_rate

                inst_completed_qty = 0.0
                inst_incomplete_qty = 0.0
                for task in inst_tasks:
                    is_done = task.state == '1_done'
                    if hasattr(task, 'item_delivery_ids'):
                        for inst_item in task.item_delivery_ids:
                            if (hasattr(inst_item, 'item_details_id') and
                                    inst_item.item_details_id and
                                    item_line.item_details_id and
                                    inst_item.item_details_id.id == item_line.item_details_id.id):
                                qty = getattr(inst_item, 'order_qty', 0) or 0.0
                                if is_done:
                                    inst_completed_qty += qty
                                else:
                                    inst_incomplete_qty += qty
                inst_completed_value = inst_completed_qty * unit_price
                inst_completed_value_incl = inst_completed_value * tax_rate
                inst_incomplete_value = inst_incomplete_qty * unit_price
                inst_incomplete_value_incl = inst_incomplete_value * tax_rate

                preview_data.append({
                    'du_workorder_no': sale.name or '',
                    'du_customer': sale.partner_id.name if sale.partner_id else '',
                    'du_customer_po_number': sale.client_order_ref or '',
                    'du_po_issue_date': sale.po_issue_date if sale.po_issue_date else False,
                    'du_order_type': order_type,
                    'du_salesperson': salesperson,
                    'du_submitted_to_customer': submitted_to_customer_str,
                    'du_submitted_date': submitted_date,
                    'du_item_name': item_line.item_name or '',
                    'du_item_description': item_line.description or '',
                    'du_workorder_qty': workorder_qty,
                    'du_unit_price': unit_price,
                    'du_total_amount': total_amount,
                    'du_total_amount_incl': tax_incl_total,
                    'du_delivered_qty': delivered_qty,
                    'du_delivered_amount': delivered_amount,
                    'du_delivered_amount_incl': delivered_amount_incl,
                    'du_undelivered_qty': undelivered_qty,
                    'du_undelivered_amount': undelivered_amount,
                    'du_undelivered_amount_incl': undelivered_amount_incl,
                    'du_inst_completed_qty': inst_completed_qty,
                    'du_inst_completed_value': inst_completed_value,
                    'du_inst_completed_value_incl': inst_completed_value_incl,
                    'du_inst_incomplete_qty': inst_incomplete_qty,
                    'du_inst_incomplete_value': inst_incomplete_value,
                    'du_inst_incomplete_value_incl': inst_incomplete_value_incl,
                })
        return preview_data

    def generate_delivered_undelivered_summary_report(self):
        """Generate Delivered/Undelivered Summary Report Excel file"""
        domain = []

        if self.du_from_date:
            domain.append(('date_order', '>=', self.du_from_date))
        if self.du_to_date:
            domain.append(('date_order', '<=', self.du_to_date))
        if self.du_customer_id:
            domain.append(('partner_id', '=', self.du_customer_id.id))
        if self.du_workorder_ids:
            domain.append(('id', 'in', self.du_workorder_ids.ids))

        sale_orders = self.env['sale.order'].search(domain, order='date_order asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Delivered Undelivered Summary')

        # ── Formats ──────────────────────────────────────────────────────────
        blue_header = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        green_header = workbook.add_format({
            'bold': True, 'bg_color': '#375623', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        red_header = workbook.add_format({
            'bold': True, 'bg_color': '#C00000', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        purple_header = workbook.add_format({
            'bold': True, 'bg_color': '#7030A0', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        orange_header = workbook.add_format({
            'bold': True, 'bg_color': '#C55A11', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        qty_format = workbook.add_format({
            'num_format': '#,##0', 'border': 1, 'valign': 'vcenter',
        })
        total_format = workbook.add_format({
            'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'valign': 'vcenter',
        })
        total_number_format = workbook.add_format({
            'bold': True, 'num_format': '#,##0.00', 'border': 1,
            'bg_color': '#D9D9D9', 'valign': 'vcenter',
        })
        total_qty_format = workbook.add_format({
            'bold': True, 'num_format': '#,##0', 'border': 1,
            'bg_color': '#D9D9D9', 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        # ── Columns ───────────────────────────────────────────────────────────
        columns = [
            ('Workorder No',                        20, blue_header),
            ('Customer PO Number',                  22, blue_header),
            ('PO Issue Date',                       15, blue_header),
            ('Customer',                            28, blue_header),
            ('Order Type',                          15, blue_header),
            ('Salesperson',                         20, blue_header),
            ('Invoice Submitted to Customer',       26, blue_header),
            ('Submitted Date',                      15, blue_header),
            ('Item Name',                           30, blue_header),
            ('Item Description',                    40, blue_header),
            ('Workorder Qty',                       14, blue_header),
            ('Unit Price',                          14, blue_header),
            ('Tax Excl. Total Amount',              20, blue_header),
            ('Tax Incl. Total Amount',              20, blue_header),
            ('Delivered Qty',                       14, blue_header),
            ('Delivered Tax Excl. Amount',          22, blue_header),
            ('Delivered Tax Incl. Amount',          22, blue_header),
            ('Undelivered Qty',                     16, blue_header),
            ('Undelivered Tax Excl. Amount',        24, blue_header),
            ('Undelivered Tax Incl. Amount',        24, blue_header),
            ('Installation Completed Qty',          22, blue_header),
            ('Installation Completed Tax Excl. Value', 28, blue_header),
            ('Installation Completed Tax Incl. Value', 28, blue_header),
            ('Installation Incomplete Qty',         22, blue_header),
            ('Installation Incomplete Tax Excl. Value', 28, blue_header),
            ('Installation Incomplete Tax Incl. Value', 28, blue_header),
        ]

        for idx, (_, width, _hfmt) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _, hfmt) in enumerate(columns):
            worksheet.write(0, col, title, hfmt)

        order_type_map = {'elcot': 'ELCOT Order', 'others': 'Direct Order'}

        grand = {
            'wo_qty': 0.0,
            'total_amt': 0.0, 'total_amt_incl': 0.0,
            'del_qty': 0.0, 'del_amt': 0.0, 'del_amt_incl': 0.0,
            'undel_qty': 0.0, 'undel_amt': 0.0, 'undel_amt_incl': 0.0,
            'inst_comp_qty': 0.0, 'inst_comp_val': 0.0, 'inst_comp_val_incl': 0.0,
            'inst_incomp_qty': 0.0, 'inst_incomp_val': 0.0, 'inst_incomp_val_incl': 0.0,
        }

        row = 1
        for sale in sale_orders:
            start_row = row
            order_type = order_type_map.get(sale.order_type, sale.order_type or '')

            done_dcs = self.env['stock.picking'].search([
                '|',
                ('sale_id', '=', sale.id),
                ('sale_order_id', '=', sale.id),
                ('picking_type_code', '=', 'outgoing'),
                ('state', '=', 'done'),
            ])
            inst_tasks = self.env['project.task'].search([('sale_order_id', '=', sale.id)])

            posted_invoices = self.env['account.move'].search([
                ('work_order_id', '=', sale.id),
                ('state', '=', 'posted'),
                ('move_type', '=', 'out_invoice'),
            ], order='invoice_date desc', limit=1)
            if posted_invoices:
                stc_val = posted_invoices.submitted_to_customer
                submitted_to_customer_str = 'Yes' if stc_val == 'yes' else ('No' if stc_val == 'no' else '')
                submitted_date = posted_invoices.submitted_date or ''
            else:
                submitted_to_customer_str = ''
                submitted_date = ''
            salesperson = sale.user_id.name if sale.user_id else ''

            if not sale.item_line_ids:
                data = [sale.name or '', sale.client_order_ref or '',
                        sale.po_issue_date if sale.po_issue_date else '',
                        sale.partner_id.name if sale.partner_id else '',
                        order_type,
                        salesperson, submitted_to_customer_str,
                        submitted_date if submitted_date else '',
                        '(No items)', '',
                        0, 0.0, 0.0, 0.0,
                        0.0, 0.0, 0.0,
                        0.0, 0.0, 0.0,
                        0.0, 0.0, 0.0,
                        0.0, 0.0, 0.0]
                for col, val in enumerate(data):
                    if col in {2, 7}:
                        fmt = date_format
                    elif col >= 10:
                        fmt = number_format
                    else:
                        fmt = cell_format
                    worksheet.write(row, col, val, fmt)
                row += 1
            else:
                for item_line in sale.item_line_ids:
                    workorder_qty = item_line.order_qty or 0.0
                    unit_price = (
                        item_line.price_subtotal / item_line.order_qty
                        if item_line.order_qty else 0.0
                    )
                    total_amount = item_line.price_subtotal or 0.0
                    tax_incl_total = item_line.price_total or 0.0
                    tax_rate = (
                        (item_line.price_total / item_line.price_subtotal)
                        if item_line.price_subtotal else 1.0
                    )

                    delivered_qty = 0.0
                    for dc in done_dcs:
                        for dc_item in dc.item_line_ids:
                            if (dc_item.item_details_id and item_line.item_details_id and
                                    dc_item.item_details_id.id == item_line.item_details_id.id):
                                delivered_qty += dc_item.order_qty or 0.0

                    delivered_amount = delivered_qty * unit_price
                    delivered_amount_incl = delivered_amount * tax_rate
                    undelivered_qty = max(0.0, workorder_qty - delivered_qty)
                    undelivered_amount = undelivered_qty * unit_price
                    undelivered_amount_incl = undelivered_qty * unit_price * tax_rate

                    inst_completed_qty = 0.0
                    inst_incomplete_qty = 0.0
                    for task in inst_tasks:
                        is_done = task.state == '1_done'
                        if hasattr(task, 'item_delivery_ids'):
                            for inst_item in task.item_delivery_ids:
                                if (hasattr(inst_item, 'item_details_id') and
                                        inst_item.item_details_id and
                                        item_line.item_details_id and
                                        inst_item.item_details_id.id == item_line.item_details_id.id):
                                    qty = getattr(inst_item, 'order_qty', 0) or 0.0
                                    if is_done:
                                        inst_completed_qty += qty
                                    else:
                                        inst_incomplete_qty += qty
                    inst_completed_value = inst_completed_qty * unit_price
                    inst_completed_value_incl = inst_completed_value * tax_rate
                    inst_incomplete_value = inst_incomplete_qty * unit_price
                    inst_incomplete_value_incl = inst_incomplete_value * tax_rate

                    grand['wo_qty'] += workorder_qty
                    grand['total_amt'] += total_amount
                    grand['total_amt_incl'] += tax_incl_total
                    grand['del_qty'] += delivered_qty
                    grand['del_amt'] += delivered_amount
                    grand['del_amt_incl'] += delivered_amount_incl
                    grand['undel_qty'] += undelivered_qty
                    grand['undel_amt'] += undelivered_amount
                    grand['undel_amt_incl'] += undelivered_amount_incl
                    grand['inst_comp_qty'] += inst_completed_qty
                    grand['inst_comp_val'] += inst_completed_value
                    grand['inst_comp_val_incl'] += inst_completed_value_incl
                    grand['inst_incomp_qty'] += inst_incomplete_qty
                    grand['inst_incomp_val'] += inst_incomplete_value
                    grand['inst_incomp_val_incl'] += inst_incomplete_value_incl

                    tax_incl_total = item_line.price_total or 0.0
                    data = [
                        sale.name or '',           # 0
                        sale.client_order_ref or '',# 1
                        sale.po_issue_date if sale.po_issue_date else '',  # 2  date
                        sale.partner_id.name if sale.partner_id else '',  # 3
                        order_type,                # 4
                        salesperson,               # 5
                        submitted_to_customer_str, # 6
                        submitted_date if submitted_date else '',  # 7  date
                        item_line.item_name or '', # 8
                        item_line.description or '',# 9
                        workorder_qty,             # 10 qty
                        unit_price,                # 11 number
                        total_amount,              # 12 number (tax excl)
                        tax_incl_total,            # 13 number (tax incl)
                        delivered_qty,             # 14 qty
                        delivered_amount,          # 15 number (tax excl)
                        delivered_amount_incl,     # 16 number (tax incl)
                        undelivered_qty,           # 17 qty
                        undelivered_amount,        # 18 number (tax excl)
                        undelivered_amount_incl,   # 19 number (tax incl)
                        inst_completed_qty,        # 20 qty
                        inst_completed_value,      # 21 number (tax excl)
                        inst_completed_value_incl, # 22 number (tax incl)
                        inst_incomplete_qty,       # 23 qty
                        inst_incomplete_value,     # 24 number (tax excl)
                        inst_incomplete_value_incl,# 25 number (tax incl)
                    ]
                    for col, val in enumerate(data):
                        if col in {2, 7}:
                            fmt = date_format
                        elif col in {10, 14, 17, 20, 23}:
                            fmt = qty_format
                        elif col in {11, 12, 13, 15, 16, 18, 19, 21, 22, 24, 25}:
                            fmt = number_format
                        else:
                            fmt = cell_format
                        worksheet.write(row, col, val, fmt)
                    row += 1

        # ── Grand Total Row ───────────────────────────────────────────────────
        total_row_data = [
            ('GRAND TOTAL',                    total_format),   # 0
            ('',                               total_format),   # 1
            ('',                               total_format),   # 2
            ('',                               total_format),   # 3
            ('',                               total_format),   # 4
            ('',                               total_format),   # 5 Salesperson
            ('',                               total_format),   # 6 Invoice Submitted
            ('',                               total_format),   # 7 Submitted Date
            ('',                               total_format),   # 8 Item Name
            ('',                               total_format),   # 9 Item Description
            (grand['wo_qty'],                  total_qty_format),   # 10
            ('',                               total_format),       # 11 Unit Price
            (grand['total_amt'],               total_number_format),# 12
            (grand['total_amt_incl'],          total_number_format),# 13
            (grand['del_qty'],                 total_qty_format),   # 14
            (grand['del_amt'],                 total_number_format),# 15
            (grand['del_amt_incl'],            total_number_format),# 16
            (grand['undel_qty'],               total_qty_format),   # 17
            (grand['undel_amt'],               total_number_format),# 18
            (grand['undel_amt_incl'],          total_number_format),# 19
            (grand['inst_comp_qty'],           total_qty_format),   # 20
            (grand['inst_comp_val'],           total_number_format),# 21
            (grand['inst_comp_val_incl'],      total_number_format),# 22
            (grand['inst_incomp_qty'],         total_qty_format),   # 23
            (grand['inst_incomp_val'],         total_number_format),# 24
            (grand['inst_incomp_val_incl'],    total_number_format),# 25
        ]
        for col, (val, fmt) in enumerate(total_row_data):
            worksheet.write(row, col, val, fmt)

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.du_from_date and self.du_to_date:
            filter_parts.append(f'{self.du_from_date}_{self.du_to_date}')
        if self.du_customer_id:
            filter_parts.append(f'customer_{self.du_customer_id.name.replace(" ", "_")}')

        filename = (
            f'delivered_undelivered_summary_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def generate_dc_print_report(self):
        """Generate DC Print Report - prints all matching DC PDFs merged into one"""
        domain = [('picking_type_code', '=', 'outgoing')]

        if self.dc_print_from_date:
            domain.append(('scheduled_date', '>=', self.dc_print_from_date))

        if self.dc_print_to_date:
            domain.append(('scheduled_date', '<=', self.dc_print_to_date))

        if self.dc_print_customer_id:
            domain.append(('partner_id', '=', self.dc_print_customer_id.id))

        if self.dc_print_warehouse_id:
            domain.append(('picking_type_id.warehouse_id', '=', self.dc_print_warehouse_id.id))

        # if self.dc_print_dc_number:
        #     domain.append(('id', '=', self.dc_print_dc_number.id))

        if self.dc_print_dc_number_ids:
            domain.append(('id', 'in', self.dc_print_dc_number_ids.ids))

        if self.dc_print_state:
            domain.append(('state', '=', self.dc_print_state))
        else:
            domain.append(('state', '=', 'done'))

        pickings = self.env['stock.picking'].search(domain, order='scheduled_date asc')

        if not pickings:
            raise ValidationError('No DC records found matching the selected filters.')

        # Build filter-based filename   soundharya
        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.dc_print_dc_number_ids:
            names = [r.dc_number or r.name for r in self.dc_print_dc_number_ids]
            parts.append(_sanitize('_'.join(names)) if len(names) <= 3 else f'{len(names)}_DCs')
        if self.dc_print_from_date:
            parts.append(self.dc_print_from_date.strftime('%d%m%Y'))
        if self.dc_print_to_date:
            parts.append('to_' + self.dc_print_to_date.strftime('%d%m%Y'))
        if self.dc_print_customer_id:
            parts.append(_sanitize(self.dc_print_customer_id.name))
        if self.dc_print_warehouse_id:
            parts.append(_sanitize(self.dc_print_warehouse_id.name, 15))
        if self.dc_print_state:
            parts.append(self.dc_print_state)
        base_name = "_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        dc_filename = f'DC_Print_{base_name}.{ext}'

        self.write({'pdf_file': False, 'pdf_filename': False})
        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        picking_ids = pickings.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._dc_print_background,
            args=(wizard_id, db, uid, picking_ids, ctx, dc_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Your DC {fmt_label} is being generated in the background.',
                'type': 'info',
                'sticky': True,
            }
        }

    def _render_dc_pdf_bytes(self, picking):
        """
        Renders a single DC PDF using the same template and data logic as the
        DC page Print button (DeliveryChallanController._prepare_dc_data).
        Returns raw PDF bytes.
        """
        import os
        import re as _re
        from jinja2 import Environment, FileSystemLoader
        from weasyprint import HTML

        # Locate template (zigma_erp/static/templates/static_dc_template.html)
        current_file = os.path.abspath(__file__)
        custom_addons = os.path.dirname(os.path.dirname(os.path.dirname(current_file)))
        template_path = os.path.join(custom_addons, 'zigma_erp', 'static', 'templates', 'static_dc_template.html')

        if not os.path.exists(template_path):
            raise Exception(f'DC template not found at {template_path}')

        jinja_env = Environment(loader=FileSystemLoader(os.path.dirname(template_path)))
        template = jinja_env.get_template(os.path.basename(template_path))

        # ── Tax totals ────────────────────────────────────────────────────────
        total_cgst = total_sgst = total_igst = 0.0
        is_igst = False
        serial = False

        for o in [picking]:
            if o.serial_needed:
                serial = True

        for item in picking.item_line_ids:
            total_cgst += item.cgst_amount or 0.0
            total_sgst += item.sgst_amount or 0.0
            total_igst += item.igst_amount or 0.0
            for tax in item.tax_id:
                if tax.tax_group_id.name == 'IGST':
                    is_igst = True
                    break
            if is_igst:
                break

        # ── Items with serial lot numbers ─────────────────────────────────────
        items_with_lots = []
        for item in picking.item_line_ids:
            moves = picking.move_ids.filtered(
                lambda mv: mv.item_details_id.id == item.item_details_id.id
            )
            lot_numbers = []
            for move in moves:
                if move.product_id.tracking != 'serial':
                    continue
                move_lines = move.move_line_ids.filtered(
                    lambda ml: ml.product_id.id == move.product_id.id
                )
                lot_numbers.extend([l for l in move_lines.mapped('lot_id.name') if l])
            items_with_lots.append({'item': item, 'lot_numbers': sorted(set(lot_numbers))})

        # ── Dates ─────────────────────────────────────────────────────────────
        inspection_date = picking.inpection_date.strftime('%d-%m-%Y') if picking.inpection_date else ''
        po_date = picking.sale_order_id.po_issue_date.strftime('%d-%m-%Y') if picking.sale_order_id and picking.sale_order_id.po_issue_date else ''
        dc_date = picking.date_done.strftime('%d-%m-%Y') if picking.date_done else ''

        # ── Partners ──────────────────────────────────────────────────────────
        consignee = picking.partner_id
        sale_order = picking.sale_order_id
        contact_person = sale_order.contact_person if sale_order else False

        if consignee.type == 'delivery':
            consignee_contact = consignee
        else:
            consignee_contact = consignee.child_ids.filtered(lambda c: c.type == 'contact')[:1]

        billing_contact = sale_order.partner_id.child_ids.filtered(lambda c: c.type == 'contact')[:1] if sale_order else self.env['res.partner']

        consignee_company = consignee.parent_id.name if consignee.parent_id else consignee.name or ''

        consignee_display_name = ''
        if consignee.comment:
            plain = _re.sub(r'<[^>]+>', ' ', consignee.comment).strip()
            plain = _re.sub(r'\s+', ' ', plain).strip()
            plain = _re.sub(r'^consignee\s*:\s*', '', plain, flags=_re.IGNORECASE).strip()
            if plain:
                consignee_display_name = plain

        consignee_name_display = consignee_display_name or consignee_company
        consignee_addr1 = consignee.street or ''
        consignee_addr2 = consignee.street2 or ''
        consignee_city_val = consignee.city or ''
        consignee_zip_val = consignee.zip or ''
        consignee_state_val = consignee.state_id.name if consignee.state_id else ''
        consignee_mobile_val = consignee_contact.mobile or consignee_contact.phone or consignee.mobile or consignee.phone or ''
        consignee_gst_val = consignee.vat or 'Unregistered'
        consignee_contact_name = consignee_contact.name or (contact_person.name if contact_person else '')

        # ── Billing section: respects tax_based_on (billing vs shipping) ──────
        tax_based_on = sale_order.tax_based_on if sale_order and hasattr(sale_order, 'tax_based_on') else 'billing'

        if tax_based_on == 'shipping':
            billing_company = consignee_name_display
            billing_addr1 = consignee_addr1
            billing_addr2 = consignee_addr2
            billing_city = consignee_city_val
            billing_zip = consignee_zip_val
            billing_state = consignee_state_val
            billing_contact_name = consignee_contact_name
            billing_mobile = consignee_mobile_val
            billing_gst = consignee_gst_val
        else:
            billing_company = sale_order.partner_id.name or '' if sale_order else ''
            billing_addr1 = sale_order.partner_id.street or '' if sale_order else ''
            billing_addr2 = sale_order.partner_id.street2 or '' if sale_order else ''
            billing_city = sale_order.partner_id.city or '' if sale_order else ''
            billing_zip = sale_order.partner_id.zip or '' if sale_order else ''
            billing_state = sale_order.partner_id.state_id.name if sale_order and sale_order.partner_id.state_id else ''
            billing_contact_name = billing_contact.name or (contact_person.name if contact_person else '')
            billing_mobile = billing_contact.mobile or billing_contact.phone or (sale_order.partner_id.mobile or sale_order.partner_id.phone or '') if sale_order else ''
            billing_gst = sale_order.partner_id.vat or 'Unregistered' if sale_order else ''

        template_data = {
            'dc_number': picking.dc_number or '',
            'gstin': picking.company_id.vat or '',
            'partner_name': consignee_contact_name,
            'contact_company': billing_company,
            'address1': billing_addr1,
            'address2': billing_addr2,
            'city': billing_city,
            'zip': billing_zip,
            'quotation_date': inspection_date,
            'po_no': sale_order.client_order_ref or '' if sale_order else '',
            'po_date': po_date,
            'items': picking.item_line_ids,
            'items_with_lots': items_with_lots,
            'total_cgst': total_cgst,
            'total_sgst': total_sgst,
            'total_igst': total_igst,
            'amount_tax': picking.amount_tax,
            'amount_total': picking.amount_total,
            'untax_total': picking.amount_untaxed,
            'amount_in_words': picking.amount_in_words,
            'is_igst': is_igst,
            'serial': serial or '',
            'dc_date': dc_date,
            'consignee_company': consignee_name_display,
            'consignee_address1': consignee_addr1,
            'consignee_address2': consignee_addr2,
            'consignee_city': consignee_city_val,
            'consignee_zip': consignee_zip_val,
            'consignee_mobile': consignee_mobile_val,
            'consignee_gst': consignee_gst_val,
            'consignee_state': consignee_state_val,
            'partner_names': billing_contact_name,
            'mobile': billing_mobile,
            'billing_state': billing_state,
            'gst': billing_gst,
            'co_no': sale_order.customer_letter or '' if sale_order else '',
            'outward_type': picking.outward_type or '',
            'outward_remark': picking.outward_remark_id.name or '' if picking.outward_remark_id else '',
        }

        html_content = template.render(template_data)
        dc_template_dir = os.path.dirname(template_path)
        return HTML(string=html_content, base_url=f'file://{dc_template_dir}/').write_pdf()

    def _dc_print_background(self, wizard_id, db, uid, picking_ids, ctx, dc_filename=None, download_type='pdf'):
        import odoo
        import io
        import zipfile

        try:
            with odoo.registry(db).cursor() as new_cr:
                env = odoo.api.Environment(new_cr, uid, ctx)
                wizard = env['reports.wizard'].browse(wizard_id)
                pickings = env['stock.picking'].browse(picking_ids)

                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for picking in pickings:
                            pdf_bytes = wizard._render_dc_pdf_bytes(picking)
                            entry_name = f"{picking.dc_number or picking.name}.pdf"
                            zf.writestr(entry_name, pdf_bytes)
                            count += 1
                    if not count:
                        raise Exception('No DC PDFs could be generated.')
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for picking in pickings:
                        pdf_bytes = wizard._render_dc_pdf_bytes(picking)
                        reader = PdfReader(io.BytesIO(pdf_bytes))
                        for page in reader.pages:
                            writer.add_page(page)
                    if not writer.pages:
                        raise Exception('No DC PDFs could be generated.')
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()

                filename = dc_filename or f'DC_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
                wizard.write({
                    'pdf_file': base64.b64encode(output_bytes).decode('utf-8'),
                    'pdf_filename': filename,
                })
                new_cr.commit()

        except Exception as e:
            _logger.error("DC print background failed: %s", str(e))

    def _get_purchase_order_preview_data(self):
        """Get preview data for Purchase Order Report"""
        domain = []

        if self.purchase_order_from_date:
            from_datetime = fields.Datetime.to_datetime(self.purchase_order_from_date)
            domain.append(('create_date_only', '>=', self.purchase_order_from_date))

        if self.purchase_order_to_date:
            domain.append(('create_date_only', '<=', self.purchase_order_to_date))

        # if self.purchase_order_id:
        #     domain.append(('id', '=', self.purchase_order_id.id))

        if self.purchase_order_ids:
            domain.append(('id', 'in', self.purchase_order_ids.ids))

        if self.purchase_order_vendor_id:
            domain.append(('partner_id', '=', self.purchase_order_vendor_id.id))

        if self.purchase_order_salesperson_id:
            domain.append(('user_id', '=', self.purchase_order_salesperson_id.id))

        if self.purchase_order_status:
            domain.append(('state', '=', self.purchase_order_status))

        purchase_orders = self.env['purchase.order'].search(domain, order='create_date asc')

        preview_data = []

        for po in purchase_orders:
            status_dict = dict(po._fields['state'].selection)
            status = status_dict.get(po.state, po.state)

            inward_pickings = self.env['stock.picking'].search([
                ('origin', '=', po.name),
                ('picking_type_code', '=', 'incoming'),
                ('state', '!=', 'cancel'),
            ], order='create_date asc')
            picking_state_labels = dict(
                self.env['stock.picking']._fields['state'].selection
            ) if inward_pickings else {}
            grn_numbers = ', '.join(p.name for p in inward_pickings) if inward_pickings else ''
            grn_dates = ', '.join(
                p.date_done.strftime('%d/%m/%Y') if p.date_done else ''
                for p in inward_pickings
            ) if inward_pickings else ''
            grn_statuses = ', '.join(
                picking_state_labels.get(p.state, p.state) for p in inward_pickings
            ) if inward_pickings else ''

            common_data = {
                'po_number': po.name or '',
                'po_vendor': po.partner_id.name if po.partner_id else '',
                'po_salesperson': po.user_id.name if po.user_id else '',
                'po_created_date': po.create_date_only if po.create_date_only else False,
                'po_expected_arrival': (po.date_planned.date() if isinstance(po.date_planned, datetime) else po.date_planned) if po.date_planned else False,
                'po_status': status,
                'po_purchase_indent': po.purchase_indent or '',
                'po_work_order_no': ', '.join(po.work_order_no.mapped('name')) if po.work_order_no else '',
                'po_customer_po_number': ', '.join(filter(None, po.work_order_no.mapped('client_order_ref'))) if po.work_order_no else '',
                'po_po_issue_date': po.work_order_no[0].po_issue_date if po.work_order_no and po.work_order_no[0].po_issue_date else False,
                'po_grn_numbers': grn_numbers,
                'po_grn_date': grn_dates,
                'po_grn_status': grn_statuses,
                'po_purchase_person': po.purchase_person_id.name if po.purchase_person_id else '',
            }

            if not po.order_line:
                row = {}
                row.update(common_data)
                row.update({
                    'po_product': '(No products)',
                    'po_description': '',
                    'po_qty_ordered': 0,
                    'po_qty_received': 0,
                    'po_qty_billed': 0,
                    'po_uom': '',
                    'po_unit_price': 0.0,
                    'po_subtotal': 0.0,
                    'po_total': 0.0,
                })
                preview_data.append(row)
                continue

            for line in po.order_line:
                row = {}
                row.update(common_data)
                row.update({
                    'po_product': line.product_id.name if line.product_id else '',
                    'po_description': line.name or '',
                    'po_qty_ordered': line.product_qty or 0,
                    'po_qty_received': line.qty_received or 0,
                    'po_qty_billed': line.qty_invoiced or 0,
                    'po_uom': line.product_uom.name if line.product_uom else '',
                    'po_unit_price': line.price_unit or 0.0,
                    'po_subtotal': line.price_subtotal or 0.0,
                    'po_total': line.price_total or 0.0,
                })
                preview_data.append(row)

        return preview_data

    def generate_purchase_order_report(self):
        """Generate Purchase Order Report Excel file"""
        domain = []

        if self.purchase_order_from_date:
            domain.append(('create_date_only', '>=', self.purchase_order_from_date))

        if self.purchase_order_to_date:
            domain.append(('create_date_only', '<=', self.purchase_order_to_date))

        # if self.purchase_order_id:
        #     domain.append(('id', '=', self.purchase_order_id.id))

        if self.purchase_order_ids:
            domain.append(('id', 'in', self.purchase_order_ids.ids))

        if self.purchase_order_vendor_id:
            domain.append(('partner_id', '=', self.purchase_order_vendor_id.id))

        if self.purchase_order_salesperson_id:
            domain.append(('user_id', '=', self.purchase_order_salesperson_id.id))

        if self.purchase_order_status:
            domain.append(('state', '=', self.purchase_order_status))

        purchase_orders = self.env['purchase.order'].search(domain, order='create_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Purchase Order Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('PO Number', 20),
            ('Vendor', 28),
            ('Salesperson', 22),
            ('Created Date', 15),
            ('Expected Arrival', 18),
            ('Status', 18),
            ('Purchase Indent', 20),
            ('Work Order No', 20),
            ('Customer PO Number', 22),
            ('PO Issue Date', 15),
            ('GRN Inward Numbers', 22),
            ('GRN Done Date', 18),
            ('GRN Inward Status', 18),
            ('Purchase Person', 22),
            ('Product', 30),
            ('Description', 35),
            ('Qty Ordered', 12),
            ('Qty Received', 12),
            ('Qty Billed', 12),
            ('UOM', 10),
            ('Unit Price', 15),
            ('Subtotal', 15),
            ('Total', 15),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        # row = 1
        # for po in purchase_orders:
        #     status_dict = dict(po._fields['state'].selection)
        #     status = status_dict.get(po.state, po.state)

        #     common = [
        #         po.name or '',
        #         po.partner_id.name if po.partner_id else '',
        #         po.user_id.name if po.user_id else '',
        #         po.create_date_only if po.create_date_only else '',
        #         po.date_planned.date() if po.date_planned else '',
        #         status,
        #         po.purchase_indent or '',
        #         po.work_order_no.name if po.work_order_no else '',
        #         po.purchase_person_id.name if po.purchase_person_id else '',
        #     ]

        #     if not po.order_line:
        #         data = common + ['(No products)', '', 0, 0, 0, '', 0.0, 0.0, 0.0]
        #         for col, val in enumerate(data):
        #             fmt = date_format if col in {3, 4} else (number_format if col >= 11 else cell_format)
        #             worksheet.write(row, col, val, fmt)
        #         row += 1
        #         continue

        #     for line in po.order_line:
        #         data = common + [
        #             line.product_id.name if line.product_id else '',
        #             line.name or '',
        #             line.product_qty or 0,
        #             line.qty_received or 0,
        #             line.qty_invoiced or 0,
        #             line.product_uom.name if line.product_uom else '',
        #             line.price_unit or 0.0,
        #             line.price_subtotal or 0.0,
        #             line.price_total or 0.0,
        #         ]
        #         for col, val in enumerate(data):
        #             if col in {3, 4}:
        #                 fmt = date_format
        #             elif col >= 11:
        #                 fmt = number_format
        #             else:
        #                 fmt = cell_format
        #             worksheet.write(row, col, val, fmt)
        #         row += 1

        row = 1
        for po in purchase_orders:
            start_row = row

            status_dict = dict(po._fields['state'].selection)
            status = status_dict.get(po.state, po.state)

            customer_po = ', '.join(filter(None, po.work_order_no.mapped('client_order_ref'))) if po.work_order_no else ''
            po_issue_date = po.work_order_no[0].po_issue_date if po.work_order_no and po.work_order_no[0].po_issue_date else ''

            inward_pickings = self.env['stock.picking'].search([
                ('origin', '=', po.name),
                ('picking_type_code', '=', 'incoming'),
                ('state', '!=', 'cancel'),
            ], order='create_date asc')
            picking_state_labels = dict(
                self.env['stock.picking']._fields['state'].selection
            ) if inward_pickings else {}
            grn_numbers = ', '.join(p.name for p in inward_pickings) if inward_pickings else ''
            grn_dates = ', '.join(
                p.date_done.strftime('%d/%m/%Y') if p.date_done else ''
                for p in inward_pickings
            ) if inward_pickings else ''
            grn_statuses = ', '.join(
                picking_state_labels.get(p.state, p.state) for p in inward_pickings
            ) if inward_pickings else ''

            common = [
                '',  # Column A - filled by merge below (PO Number)
                po.partner_id.name if po.partner_id else '',
                po.user_id.name if po.user_id else '',
                po.create_date_only if po.create_date_only else '',
                (po.date_planned.date() if isinstance(po.date_planned, datetime) else po.date_planned) if po.date_planned else '',
                status,
                po.purchase_indent or '',
                ', '.join(po.work_order_no.mapped('name')) if po.work_order_no else '',
                '',  # Customer PO Number - filled by merge below
                '',  # PO Issue Date - filled by merge below
                '',  # GRN Inward Numbers - filled by merge below
                '',  # GRN Created Date - filled by merge below
                '',  # GRN Inward Status - filled by merge below
                po.purchase_person_id.name if po.purchase_person_id else '',
            ]

            if not po.order_line:
                data = common + ['(No products)', '', 0, 0, 0, '', 0.0, 0.0, 0.0]
                for col, val in enumerate(data):
                    if col in {3, 4, 9}:
                        fmt = date_format
                    elif col >= 16:
                        fmt = number_format
                    else:
                        fmt = cell_format
                    worksheet.write(row, col, val, fmt)
                row += 1
            else:
                for line in po.order_line:
                    data = common + [
                        line.product_id.name if line.product_id else '',
                        line.name or '',
                        line.product_qty or 0,
                        line.qty_received or 0,
                        line.qty_invoiced or 0,
                        line.product_uom.name if line.product_uom else '',
                        line.price_unit or 0.0,
                        line.price_subtotal or 0.0,
                        line.price_total or 0.0,
                    ]
                    for col, val in enumerate(data):
                        if col in {3, 4, 9}:
                            fmt = date_format
                        elif col >= 16:
                            fmt = number_format
                        else:
                            fmt = cell_format
                        worksheet.write(row, col, val, fmt)
                    row += 1

            end_row = row - 1

            # Merge PO Number (col 0), Customer PO Number (col 8), PO Issue Date (col 9),
            # GRN Inward Numbers (col 10), GRN Created Date (col 11), GRN Inward Status (col 12)
            if end_row > start_row:
                worksheet.merge_range(start_row, 0, end_row, 0, po.name or '', cell_format)
                worksheet.merge_range(start_row, 8, end_row, 8, customer_po, cell_format)
                worksheet.merge_range(start_row, 9, end_row, 9, po_issue_date, date_format if po_issue_date else cell_format)
                worksheet.merge_range(start_row, 10, end_row, 10, grn_numbers, cell_format)
                worksheet.merge_range(start_row, 11, end_row, 11, grn_dates, cell_format)
                worksheet.merge_range(start_row, 12, end_row, 12, grn_statuses, cell_format)
            else:
                worksheet.write(start_row, 0, po.name or '', cell_format)
                worksheet.write(start_row, 8, customer_po, cell_format)
                worksheet.write(start_row, 9, po_issue_date, date_format if po_issue_date else cell_format)
                worksheet.write(start_row, 10, grn_numbers, cell_format)
                worksheet.write(start_row, 11, grn_dates, cell_format)
                worksheet.write(start_row, 12, grn_statuses, cell_format)

        workbook.close()
        output.seek(0)

        filter_parts = []
        # if self.purchase_order_id:
        #     filter_parts.append(f'po_{self.purchase_order_id.name.replace("/", "_")}')
        if self.purchase_order_ids:
            po_names = "_".join(self.purchase_order_ids.mapped('name'))
            filter_parts.append(f'po_{po_names.replace("/", "_")}')
        if self.purchase_order_from_date and self.purchase_order_to_date:
            filter_parts.append(f'{self.purchase_order_from_date}_{self.purchase_order_to_date}')
        if self.purchase_order_vendor_id:
            filter_parts.append(f'vendor_{self.purchase_order_vendor_id.name.replace(" ", "_")}')

        filename = (
            f'purchase_order_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_purchase_indent_preview_data(self):
        """Get preview data for Purchase Indent Report"""
        domain = []

        # if self.purchase_indent_id:
        #     domain.append(('id', '=', self.purchase_indent_id.id))

        if self.purchase_indent_ids:
            domain.append(('id', 'in', self.purchase_indent_ids.ids))

        if self.purchase_indent_from_date:
            domain.append(('date', '>=', self.purchase_indent_from_date))

        if self.purchase_indent_to_date:
            domain.append(('date', '<=', self.purchase_indent_to_date))

        if self.purchase_indent_vendor_id:
            domain.append(('vendor_id', '=', self.purchase_indent_vendor_id.id))

        if self.purchase_indent_order_type:
            domain.append(('order_type', '=', self.purchase_indent_order_type))

        if self.purchase_indent_status:
            domain.append(('state', '=', self.purchase_indent_status))

        indents = self.env['purchase.indent'].search(domain, order='date asc')

        preview_data = []

        for indent in indents:
            status_dict = dict(indent._fields['state'].selection)
            status = status_dict.get(indent.state, indent.state)

            order_type_map = {
                'workorder': 'Against Work order',
                'direct': 'Direct',
            }
            order_type = order_type_map.get(indent.order_type, indent.order_type or '')

            common_data = {
                'pi_indent_no': indent.name or '',
                'pi_indent_date': indent.date.date() if indent.date else False,
                'pi_order_type': order_type,
                'pi_sale_order': indent.sale_order_id.name if indent.sale_order_id else '',
                'pi_approved_date': indent.approved_date.date() if indent.approved_date else False,
                'pi_vendor': indent.vendor_id.name if indent.vendor_id else '',
                'pi_status': status,
            }

            if not indent.line_ids:
                row = {}
                row.update(common_data)
                row.update({
                    'pi_product': '(No products)',
                    'pi_part_number': '',
                    'pi_qty': 0.0,
                    'pi_price': 0.0,
                    'pi_uom': '',
                })
                preview_data.append(row)
                continue

            for line in indent.line_ids:
                row = {}
                row.update(common_data)
                row.update({
                    'pi_product': line.product_id.name if line.product_id else '',
                    'pi_part_number': line.part_number or '',
                    'pi_qty': line.qty or 0.0,
                    'pi_price': line.price or 0.0,
                    'pi_uom': line.uom_id.name if line.uom_id else '',
                })
                preview_data.append(row)

        return preview_data

    def generate_purchase_indent_report(self):
        """Generate Purchase Indent Report Excel file"""
        domain = []

        # if self.purchase_indent_id:
        #     domain.append(('id', '=', self.purchase_indent_id.id))

        if self.purchase_indent_ids:
            domain.append(('id', 'in', self.purchase_indent_ids.ids))

        if self.purchase_indent_from_date:
            domain.append(('date', '>=', self.purchase_indent_from_date))

        if self.purchase_indent_to_date:
            domain.append(('date', '<=', self.purchase_indent_to_date))

        if self.purchase_indent_vendor_id:
            domain.append(('vendor_id', '=', self.purchase_indent_vendor_id.id))

        if self.purchase_indent_order_type:
            domain.append(('order_type', '=', self.purchase_indent_order_type))

        if self.purchase_indent_status:
            domain.append(('state', '=', self.purchase_indent_status))

        indents = self.env['purchase.indent'].search(domain, order='date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Purchase Indent Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Indent No',            20),
            ('Indent Date',          15),
            ('Order Type',           20),
            ('Sale Order',           20),
            ('Indent Approved Date', 20),
            ('Vendor',               25),
            ('Status',               18),
            ('Product',              35),
            ('Part Number',          20),
            ('Qty',                  10),
            ('Price',                15),
            ('UOM',                  12),
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        order_type_map = {
            'workorder': 'Against Work order',
            'direct': 'Direct',
        }

        row = 1
        for indent in indents:
            start_row = row
            status_dict = dict(indent._fields['state'].selection)
            status = status_dict.get(indent.state, indent.state)
            order_type = order_type_map.get(indent.order_type, indent.order_type or '')

            common = [
                indent.name or '',
                indent.date.date() if indent.date else '',
                order_type,
                indent.sale_order_id.name if indent.sale_order_id else '',
                indent.approved_date.date() if indent.approved_date else '',
                indent.vendor_id.name if indent.vendor_id else '',
                status,
            ]

            if not indent.line_ids:
                data = common + ['(No products)', '', 0.0, 0.0, '']
                for col, val in enumerate(data):
                    fmt = date_format if col in {1, 4} else (number_format if col in {9, 10} else cell_format)
                    worksheet.write(row, col, val, fmt)
                row += 1
                continue

            for line in indent.line_ids:
                data = common + [
                    line.product_id.name if line.product_id else '',
                    line.part_number or '',
                    line.qty or 0.0,
                    line.price or 0.0,
                    line.uom_id.name if line.uom_id else '',
                ]
                for col, val in enumerate(data):
                    if col in {1, 4}:
                        fmt = date_format
                    elif col in {9, 10}:
                        fmt = number_format
                    else:
                        fmt = cell_format
                    worksheet.write(row, col, val, fmt)
                row += 1

            end_row = row - 1

            if start_row == end_row:
                worksheet.write(start_row, 0, indent.name or '', cell_format)
            else:
                worksheet.merge_range(start_row, 0, end_row, 0, indent.name or '', cell_format)

        workbook.close()
        output.seek(0)

        filter_parts = []
        # if self.purchase_indent_id:
        #     filter_parts.append(f'indent_{self.purchase_indent_id.name.replace("/", "_")}')

        if self.purchase_indent_ids:
            names = "_".join(self.purchase_indent_ids.mapped('name'))
            filter_parts.append(f'indent_{names.replace("/", "_")}')
        if self.purchase_indent_from_date and self.purchase_indent_to_date:
            filter_parts.append(f'{self.purchase_indent_from_date}_{self.purchase_indent_to_date}')
        if self.purchase_indent_vendor_id:
            filter_parts.append(f'vendor_{self.purchase_indent_vendor_id.name.replace(" ", "_")}')

        filename = (
            f'purchase_indent_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    def _get_invoice_summary_preview_data(self):
        """Get preview data for Invoice Summary Report"""
        domain = [('move_type', '=', 'out_invoice')]

        # if self.invoice_summary_id:
        #     domain.append(('id', '=', self.invoice_summary_id.id))

        if self.invoice_summary_ids:
            domain.append(('id', 'in', self.invoice_summary_ids.ids))

        if self.invoice_summary_from_date:
            domain.append(('invoice_date', '>=', self.invoice_summary_from_date))

        if self.invoice_summary_to_date:
            domain.append(('invoice_date', '<=', self.invoice_summary_to_date))

        if self.invoice_summary_customer_id:
            domain.append(('partner_id', '=', self.invoice_summary_customer_id.id))

        if self.invoice_summary_invoice_type:
            domain.append(('invoice_type', '=', self.invoice_summary_invoice_type))

        if self.invoice_summary_status:
            domain.append(('state', '=', self.invoice_summary_status))

        invoices = self.env['account.move'].search(domain, order='invoice_date asc')

        preview_data = []

        for inv in invoices:
            status_dict = dict(inv._fields['state'].selection)
            status = status_dict.get(inv.state, inv.state)

            invoice_type_map = {
                'customer': 'Customer',
                'consignee': 'Consignee',
            }
            inv_type = invoice_type_map.get(inv.invoice_type, inv.invoice_type or '') if hasattr(inv, 'invoice_type') else ''

            # Get DC number from selected_dc_id
            dc_number = ''
            if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
                dc_number = inv.selected_dc_id.dc_number or inv.selected_dc_id.name or ''

            # Get work order
            work_order = ''
            order_type_val = ''
            if hasattr(inv, 'work_order_id') and inv.work_order_id:
                work_order = inv.work_order_id.name or ''
                ot = inv.work_order_id.order_type or ''
                order_type_field = inv.work_order_id._fields.get('order_type')
                if order_type_field:
                    order_type_val = dict(order_type_field.selection).get(ot, ot)
                else:
                    order_type_val = ot

            district_val = inv.partner_id.city or '' if inv.partner_id else ''

            common_data = {
                'inv_number': inv.name or '',
                'inv_customer': inv.invoice_partner_display_name or (inv.partner_id.name if inv.partner_id else ''),
                'inv_district': district_val,
                'inv_date': inv.invoice_date if inv.invoice_date else False,
                'inv_due_date': inv.invoice_date_due if inv.invoice_date_due else False,
                'inv_dc_number': dc_number,
                'inv_work_order': work_order,
                'inv_order_type': order_type_val,
                'inv_po_number': inv.po_number or '' if hasattr(inv, 'po_number') else '',
                'inv_po_issue_date': inv.work_order_id.po_issue_date if hasattr(inv, 'work_order_id') and inv.work_order_id and inv.work_order_id.po_issue_date else False,
                'inv_invoice_type': inv_type,
                'inv_status': status,
            }

            # Get item lines from stock.picking.item.line via dc
            item_lines = []
            if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
                if hasattr(inv.selected_dc_id, 'item_line_ids'):
                    item_lines = inv.selected_dc_id.item_line_ids

            if item_lines:
                for line in item_lines:
                    taxes_str = ', '.join(
                        re.sub(r'\s+[SP]$', '', n, flags=re.IGNORECASE).strip()
                        for n in line.tax_id.mapped('name')
                    ) if line.tax_id else ''
                    row = {}
                    row.update(common_data)
                    row.update({
                        'inv_item_name': line.item_name or '',
                        'inv_description': line.description or '',
                        'inv_qty': line.order_qty or 0,
                        'inv_taxes': taxes_str,
                        'inv_subtotal': line.price_subtotal or 0.0,
                    })
                    preview_data.append(row)
            else:
                row = {}
                row.update(common_data)
                row.update({
                    'inv_item': '',
                    'inv_item_name': '',
                    'inv_description': '',
                    'inv_qty': 0,
                    'inv_taxes': '',
                    'inv_subtotal': inv.amount_untaxed or 0.0,
                })
                preview_data.append(row)

        return preview_data

    def generate_invoice_summary_report(self):
        """Generate Invoice Summary Report Excel file"""
        domain = [('move_type', '=', 'out_invoice')]

        # if self.invoice_summary_id:
        #     domain.append(('id', '=', self.invoice_summary_id.id))

        if self.invoice_summary_ids:
            domain.append(('id', 'in', self.invoice_summary_ids.ids))

        if self.invoice_summary_from_date:
            domain.append(('invoice_date', '>=', self.invoice_summary_from_date))

        if self.invoice_summary_to_date:
            domain.append(('invoice_date', '<=', self.invoice_summary_to_date))

        if self.invoice_summary_customer_id:
            domain.append(('partner_id', '=', self.invoice_summary_customer_id.id))

        if self.invoice_summary_invoice_type:
            domain.append(('invoice_type', '=', self.invoice_summary_invoice_type))

        if self.invoice_summary_status:
            domain.append(('state', '=', self.invoice_summary_status))

        invoices = self.env['account.move'].search(domain, order='invoice_date asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Invoice Summary Report')

        header_format = workbook.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00', 'border': 1, 'valign': 'vcenter',
        })
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy', 'border': 1, 'valign': 'vcenter',
        })

        columns = [
            ('Invoice Number',        22),  # 0
            ('Customer',              30),  # 1
            ('District',              20),  # 2
            ('Invoice Date',          15),  # 3  date
            ('Due Date',              15),  # 4  date
            ('DC Number',             20),  # 5
            ('Work Order',            30),  # 6
            ('Order Type',            20),  # 7
            ('Customer PO Number',    30),  # 8
            ('PO Issue Date',         15),  # 9  date
            ('Salesperson',           20),  # 10
            ('Submitted to Customer', 20),  # 11
            ('Submitted Date',        15),  # 12 date
            ('Invoice Type',          15),  # 13
            ('GST Number',            25),  # 14
            ('Status',                15),  # 15
            ('Item Name',             25),  # 16
            ('Description',           40),  # 17
            ('Quantity',              12),  # 18 number
            ('Unit Price',            15),  # 19 number
            ('Taxes',                 20),  # 20
            ('SGST Amount',           15),  # 21 number
            ('CGST Amount',           15),  # 22 number
            ('IGST Amount',           15),  # 23 number
            ('Tax Amount',            15),  # 24 number
            ('Tax Excluded',          15),  # 25 number
            ('Total Amount',          15),  # 26 number
            ('Grand Total',           15),  # 27 number
        ]

        for idx, (_, width) in enumerate(columns):
            worksheet.set_column(idx, idx, width)
        for col, (title, _) in enumerate(columns):
            worksheet.write(0, col, title, header_format)

        invoice_type_map = {
            'customer': 'Customer',
            'consignee': 'Consignee',
        }

        # row = 1
        # for inv in invoices:
        #     status_dict = dict(inv._fields['state'].selection)
        #     status = status_dict.get(inv.state, inv.state)
        #     inv_type = invoice_type_map.get(inv.invoice_type, inv.invoice_type or '') if hasattr(inv, 'invoice_type') else ''

        #     dc_number = ''
        #     if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
        #         dc_number = inv.selected_dc_id.dc_number or inv.selected_dc_id.name or ''

        #     work_order = ''
        #     if hasattr(inv, 'work_order_id') and inv.work_order_id:
        #         work_order = inv.work_order_id.name or ''

        #     common = [
        #         inv.name or '',
        #         inv.invoice_partner_display_name or (inv.partner_id.name if inv.partner_id else ''),
        #         inv.invoice_date if inv.invoice_date else '',
        #         inv.invoice_date_due if inv.invoice_date_due else '',
        #         dc_number,
        #         work_order,
        #         inv.po_number or '' if hasattr(inv, 'po_number') else '',
        #         inv_type,
        #         status,
        #     ]

        #     item_lines = []
        #     if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
        #         if hasattr(inv.selected_dc_id, 'item_line_ids'):
        #             item_lines = inv.selected_dc_id.item_line_ids

        #     if item_lines:
        #         for line in item_lines:
        #             taxes_str = ', '.join(line.tax_id.mapped('name')) if line.tax_id else ''
        #             data = common + [
        #                 line.item_details_id.name if line.item_details_id else '',
        #                 line.item_name or '',
        #                 line.description or '',
        #                 line.order_qty or 0,
        #                 taxes_str,
        #                 line.price_subtotal or 0.0,
        #             ]
        #             for col, val in enumerate(data):
        #                 if col in {2, 3}:
        #                     fmt = date_format
        #                 elif col in {12, 14}:
        #                     fmt = number_format
        #                 else:
        #                     fmt = cell_format
        #                 worksheet.write(row, col, val, fmt)
        #             row += 1
        #     else:
        #         data = common + ['', '', '', 0, '', inv.amount_untaxed or 0.0]
        #         for col, val in enumerate(data):
        #             if col in {2, 3}:
        #                 fmt = date_format
        #             elif col in {12, 14}:
        #                 fmt = number_format
        #             else:
        #                 fmt = cell_format
        #             worksheet.write(row, col, val, fmt)
        #         row += 1

        row = 1
        for inv in invoices:
            status_dict = dict(inv._fields['state'].selection)
            status = status_dict.get(inv.state, inv.state)
            inv_type = invoice_type_map.get(inv.invoice_type, inv.invoice_type or '') if hasattr(inv, 'invoice_type') else ''

            dc_number = ''
            if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
                dc_number = inv.selected_dc_id.dc_number or inv.selected_dc_id.name or ''

            work_order = ''
            order_type_val = ''
            if hasattr(inv, 'work_order_id') and inv.work_order_id:
                work_order = inv.work_order_id.name or ''
                ot = inv.work_order_id.order_type or ''
                order_type_field = inv.work_order_id._fields.get('order_type')
                if order_type_field:
                    order_type_val = dict(order_type_field.selection).get(ot, ot)
                else:
                    order_type_val = ot

            submitted_to_cust_map = {'yes': 'Yes', 'no': 'No'}
            submitted_to_cust_str = submitted_to_cust_map.get(
                inv.submitted_to_customer if hasattr(inv, 'submitted_to_customer') else '', '')
            submitted_date_val = inv.submitted_date if hasattr(inv, 'submitted_date') and inv.submitted_date else ''
            salesperson_val = (
                inv.work_order_id.user_id.name
                if hasattr(inv, 'work_order_id') and inv.work_order_id and inv.work_order_id.user_id
                else inv.invoice_user_id.name if inv.invoice_user_id else ''
            )
            grand_total_val = inv.item_total_amount if hasattr(inv, 'item_total_amount') and inv.item_total_amount else inv.amount_total or 0.0

            if hasattr(inv, 'invoice_type') and inv.invoice_type == 'consignee':
                gst_number_val = (inv.consignee_address_id.vat or '') if hasattr(inv, 'consignee_address_id') and inv.consignee_address_id else ''
            else:
                gst_number_val = (inv.partner_id.vat or '') if inv.partner_id else ''

            district_val = inv.partner_id.city or '' if inv.partner_id else ''

            common = [
                inv.name or '',                                                                    # 0
                inv.invoice_partner_display_name or (inv.partner_id.name if inv.partner_id else ''),  # 1
                district_val,                                                                      # 2
                inv.invoice_date if inv.invoice_date else '',                                      # 3
                inv.invoice_date_due if inv.invoice_date_due else '',                              # 4
                dc_number,                                                                         # 5
                work_order,                                                                        # 6
                order_type_val,                                                                    # 7
                inv.po_number or '' if hasattr(inv, 'po_number') else '',                         # 8
                inv.work_order_id.po_issue_date if hasattr(inv, 'work_order_id') and inv.work_order_id and inv.work_order_id.po_issue_date else '',  # 9
                salesperson_val,                                                                   # 10
                submitted_to_cust_str,                                                             # 11
                submitted_date_val,                                                                # 12
                inv_type,                                                                          # 13
                gst_number_val,                                                                    # 14
                status,                                                                            # 15
            ]

            # item_lines = []
            # if hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
            #     if hasattr(inv.selected_dc_id, 'item_line_ids'):
            #         item_lines = inv.selected_dc_id.item_line_ids

            item_lines = []

            # DEBUG: Log all fields related to DC on this invoice
            import logging
            _logger = logging.getLogger(__name__)
            dc_related_fields = [f for f in inv._fields.keys() if 'dc' in f.lower() or 'delivery' in f.lower() or 'picking' in f.lower()]
            _logger.warning('Invoice %s - DC related fields: %s', inv.name, dc_related_fields)

            # Try to find the correct many2many DC field
            for field_name in ['dc_ids', 'delivery_ids', 'picking_ids', 'stock_picking_ids', 
                            'related_dc_ids', 'dc_id', 'selected_dc_ids']:
                if hasattr(inv, field_name):
                    field_val = getattr(inv, field_name)
                    _logger.warning('Invoice %s - Field %s = %s', inv.name, field_name, field_val)
                    if field_val:
                        # It's a many2many/one2many
                        if hasattr(field_val, '__iter__') and not isinstance(field_val, str):
                            for dc in field_val:
                                if hasattr(dc, 'item_line_ids') and dc.item_line_ids:
                                    item_lines += list(dc.item_line_ids)
                                    _logger.warning('Invoice %s - Found %s item lines from field %s dc %s', 
                                                inv.name, len(dc.item_line_ids), field_name, dc.name)
                        break

            # Fallback to single selected_dc_id
            if not item_lines and hasattr(inv, 'selected_dc_id') and inv.selected_dc_id:
                if hasattr(inv.selected_dc_id, 'item_line_ids'):
                    item_lines = list(inv.selected_dc_id.item_line_ids)

            # Final fallback
            if not item_lines and inv.invoice_line_ids:
                item_lines = list(inv.invoice_line_ids)


            if item_lines:
                start_row = row
                for line in item_lines:
                    # DC item lines use tax_id; standard account.move.line uses tax_ids
                    if hasattr(line, 'tax_id') and line.tax_id:
                        taxes_str = ', '.join(
                            re.sub(r'\s+[SP]$', '', n, flags=re.IGNORECASE).strip()
                            for n in line.tax_id.mapped('name')
                        )
                    elif hasattr(line, 'tax_ids') and line.tax_ids:
                        taxes_str = ', '.join(
                            re.sub(r'\s+[SP]$', '', n, flags=re.IGNORECASE).strip()
                            for n in line.tax_ids.mapped('name')
                        )
                    else:
                        taxes_str = ''
                    # DC item lines have custom fields; account.move.line uses standard fields
                    item_code = line.item_details_id.name if hasattr(line, 'item_details_id') and line.item_details_id else ''
                    item_name = line.item_name if hasattr(line, 'item_name') and line.item_name else (line.product_id.name if hasattr(line, 'product_id') and line.product_id else '')
                    description = line.description if hasattr(line, 'description') and line.description else (line.name if hasattr(line, 'name') and line.name else '')
                    qty = line.order_qty if hasattr(line, 'order_qty') and line.order_qty else (line.quantity if hasattr(line, 'quantity') else 0)
                    unit_price = line.price_unit if hasattr(line, 'price_unit') and line.price_unit else 0.0
                    igst = line.igst_amount if hasattr(line, 'igst_amount') and line.igst_amount else 0.0
                    if igst:
                        sgst = 0.0
                        cgst = 0.0
                    else:
                        sgst = line.sgst_amount if hasattr(line, 'sgst_amount') and line.sgst_amount else 0.0
                        cgst = line.cgst_amount if hasattr(line, 'cgst_amount') and line.cgst_amount else 0.0
                    total_amount = line.price_total if hasattr(line, 'price_total') and line.price_total else (line.price_subtotal or 0.0) + sgst + cgst + igst
                    data = common + [
                        item_name,       # 16
                        description,     # 17
                        qty,             # 18
                        unit_price,      # 19
                        taxes_str,       # 20
                        sgst,            # 21
                        cgst,            # 22
                        igst,            # 23
                        sgst + cgst + igst,        # 24
                        line.price_subtotal or 0.0, # 25
                        total_amount,    # 26
                    ]
                    for col, val in enumerate(data):
                        if col in {3, 4, 9, 12}:
                            fmt = date_format
                        elif col in {18, 19, 21, 22, 23, 24, 25, 26}:
                            fmt = number_format
                        else:
                            fmt = cell_format
                        worksheet.write(row, col, val, fmt)
                    row += 1
                end_row = row - 1
                if end_row > start_row:
                    worksheet.merge_range(start_row, 27, end_row, 27, grand_total_val, number_format)
                else:
                    worksheet.write(start_row, 27, grand_total_val, number_format)
            else:
                inv_igst = inv.item_igst_amount if hasattr(inv, 'item_igst_amount') and inv.item_igst_amount else 0.0
                if inv_igst:
                    inv_sgst = 0.0
                    inv_cgst = 0.0
                else:
                    inv_sgst = inv.item_sgst_amount if hasattr(inv, 'item_sgst_amount') and inv.item_sgst_amount else 0.0
                    inv_cgst = inv.item_cgst_amount if hasattr(inv, 'item_cgst_amount') and inv.item_cgst_amount else 0.0
                inv_total = inv.item_total_amount if hasattr(inv, 'item_total_amount') and inv.item_total_amount else inv.amount_total or 0.0
                data = common + ['', '', 0, 0.0, '', inv_sgst, inv_cgst, inv_igst, inv_sgst + inv_cgst + inv_igst, inv.amount_untaxed or 0.0, inv_total]
                for col, val in enumerate(data):
                    if col in {3, 4, 9, 12}:
                        fmt = date_format
                    elif col in {18, 19, 21, 22, 23, 24, 25, 26}:
                        fmt = number_format
                    else:
                        fmt = cell_format
                    worksheet.write(row, col, val, fmt)
                worksheet.write(row, 27, grand_total_val, number_format)
                row += 1

        workbook.close()
        output.seek(0)

        filter_parts = []
        if self.invoice_summary_id:
            filter_parts.append(f'inv_{self.invoice_summary_id.name.replace("/", "_")}')
        if self.invoice_summary_from_date and self.invoice_summary_to_date:
            filter_parts.append(f'{self.invoice_summary_from_date}_{self.invoice_summary_to_date}')
        if self.invoice_summary_customer_id:
            filter_parts.append(f'customer_{self.invoice_summary_customer_id.name.replace(" ", "_")}')

        filename = (
            f'invoice_summary_report_'
            f'{"_".join(filter_parts) if filter_parts else "all"}_'
            f'{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        self.write({
            'excel_file': base64.b64encode(output.read()),
            'filename': filename,
        })

        self.generate_preview()

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'view_mode': 'form',
            'res_id': self.id,
            'target': 'current',
            'views': [(view_id, 'form')],
        }

    # ============================================================
    # ADD this helper method to ReportsWizard class
    # ============================================================

    def _render_invoice_pdf_bytes(self, invoice):
        """
        Renders the Tax Invoice PDF bytes directly by replicating
        the logic in PurchaseChallanController.generate_dc_pdf()
        without making an HTTP request.
        """
        import os
        from jinja2 import Environment, FileSystemLoader
        from weasyprint import HTML

        # ── Tax calculations (same as controller) ─────────────────
        currency = invoice.currency_id or self.env.company.currency_id
        total_amount = invoice.item_total_amount or 0.0
        amount_in_words = currency.amount_to_text(total_amount)
        amount_in_words = amount_in_words.replace('-', ' ').replace(',', '') + ' Only'

        fiscal_position = invoice.fiscal_position_id
        total_cgst = total_sgst = total_igst = 0.0
        is_igst = False

        for item in invoice.item_delivery_ids:
            total_cgst += item.cgst_amount or 0.0
            total_sgst += item.sgst_amount or 0.0
            total_igst += item.igst_amount or 0.0
            taxes = item.tax_id
            mapped_taxes = fiscal_position.map_tax(taxes) if fiscal_position else taxes
            for tax in mapped_taxes:
                if tax.tax_group_id.name == 'IGST':
                    is_igst = True
                    break
            if is_igst:
                break

        if not total_cgst and not total_sgst and not total_igst:
            total_cgst = invoice.item_cgst_amount or 0.0
            total_sgst = invoice.item_sgst_amount or 0.0
            total_igst = invoice.item_igst_amount or 0.0

        # ── Build items_with_lots (same as controller) ─────────────
        items_with_lots = []
        for item in invoice.item_delivery_ids:
            item_data = {'item': item, 'lot_numbers': []}
            if invoice.selected_dc_id and invoice.selected_dc_id.move_ids:
                moves = invoice.selected_dc_id.move_ids.filtered(
                    lambda mv: mv.item_details_id.id == item.item_details_id.id
                )
                lot_numbers = []
                for move in moves:
                    if move.product_id.tracking != 'serial':
                        continue
                    if move.item_lines:
                        lot_names = [l.lot_id.name for l in move.item_lines if l.lot_id and not l.lot_id.auto_serial_generated]
                    else:
                        ml_filtered = move.move_line_ids.filtered(lambda ml: ml.product_id.id == move.product_id.id)
                        lot_names = [ml.lot_id.name for ml in ml_filtered if ml.lot_id and not ml.lot_id.auto_serial_generated]
                    lot_numbers.extend(lot_names)
                item_data['lot_numbers'] = list(set(lot_numbers))
            items_with_lots.append(item_data)

        # ── Address logic matching accounting/controllers/pdf_api.py ─────────────
        # Consignee: use consignee_address_id if set, else partner_id
        consignee_partner = invoice.consignee_address_id if invoice.consignee_address_id else invoice.partner_id
        consignee_display = consignee_partner.parent_id.name if consignee_partner.parent_id else consignee_partner.name or ''
        consignee_person = consignee_partner.name or ''
        consignee_contact = consignee_partner.child_ids.filtered(lambda c: c.type == 'contact')[:1]
        if consignee_contact:
            consignee_person = consignee_contact.name or consignee_person
            consignee_phone = consignee_contact.phone or consignee_contact.mobile or ''
        else:
            consignee_phone = consignee_partner.phone or consignee_partner.mobile or ''

        # Billing: if invoice_type == 'consignee' and consignee_address_id set, use consignee; else use partner_id
        if invoice.invoice_type == 'consignee' and invoice.consignee_address_id:
            billing_partner = invoice.consignee_address_id
        else:
            billing_partner = invoice.partner_id
        billing_name    = billing_partner.parent_id.name if billing_partner.parent_id else billing_partner.name or ''
        billing_add     = billing_partner.street or ''
        billing_add2    = billing_partner.street2 or ''
        billing_city    = billing_partner.city or ''
        billing_zip     = billing_partner.zip or ''
        billing_state   = billing_partner.state_id.name if billing_partner.state_id else ''
        billing_country = billing_partner.country_id.name if billing_partner.country_id else ''
        bil_vat         = billing_partner.vat or 'UNREGISTERED'

        # ── FIX: acco.html is in the 'accounting' module, not 'hrms_dashboard' ──
        # reports_wizard.py is in: hrms_dashboard/models/
        # acco.html is in:         accounting/static/templates/
        # Both modules are under the same custom-addons directory.

        current_file    = os.path.abspath(__file__)           # hrms_dashboard/models/reports_wizard.py
        models_dir      = os.path.dirname(current_file)       # hrms_dashboard/models/
        hrms_root       = os.path.dirname(models_dir)         # hrms_dashboard/
        custom_addons   = os.path.dirname(hrms_root)          # custom-addons/
        
        template_dir  = os.path.join(custom_addons, 'accounting', 'static', 'templates')
        template_name = 'acco.html'
        html_path     = os.path.join(template_dir, template_name)

        # Verify the file actually exists
        if not os.path.exists(html_path):
            raise ValidationError(
                f'Invoice template not found at: {html_path}\n'
                f'Expected: custom-addons/accounting/static/templates/acco.html'
            )

        env_j2   = Environment(loader=FileSystemLoader(template_dir))
        template = env_j2.get_template(template_name)

        html_content = template.render({
            'quotation_no':        invoice.name or '',
            'gstin':               invoice.partner_id.vat or '',
            'consignee':           consignee_person,
            'contact_company':     consignee_display,
            'address1':            consignee_partner.street or '',
            'address2':            consignee_partner.street2 or '',
            'city':                consignee_partner.city or '',
            'zip':                 consignee_partner.zip or '',
            'state':               consignee_partner.state_id.name if consignee_partner.state_id else '',
            'country':             consignee_partner.country_id.name if consignee_partner.country_id else '',
            'consignee_gstin':     consignee_partner.vat or '',
            'quotation_date':      invoice.invoice_date.strftime('%d-%m-%Y') if invoice.invoice_date else '',
            'po_no':               invoice.po_number or '',
            'po_date':             invoice.po_date.strftime('%d-%m-%Y') if invoice.po_date else '',
            'phone':               consignee_phone,
            'items_with_lots':     items_with_lots,
            'dc_no':               invoice.selected_dc_id.dc_number or '',
            'total_cgst':          total_cgst,
            'total_sgst':          total_sgst,
            'total_igst':          total_igst,
            'amount_tax':          invoice.item_tax_amount or 0.0,
            'enquiry_no':          invoice.ref or '',
            'amount_totals':       total_amount,
            'item_untaxed_amount': invoice.item_untaxed_amount or 0.0,
            'date_order':          invoice.invoice_date.strftime('%d-%m-%Y') if invoice.invoice_date else '',
            'amount_in_words':     amount_in_words,
            'is_igst':             is_igst,
            'work_order':          invoice.work_order_id.name if invoice.work_order_id else '',
            'delivery_date':       invoice.delivery_date.strftime('%d-%m-%Y') if hasattr(invoice, 'delivery_date') and invoice.delivery_date else '',
            'due_date':            invoice.selected_dc_id.scheduled_date.strftime('%d-%m-%Y') if invoice.selected_dc_id and invoice.selected_dc_id.scheduled_date else '',
            'payment_terms':       invoice.invoice_payment_term_id.name if invoice.invoice_payment_term_id else '',
            'origin':              invoice.invoice_origin or '',
            'billing_address':     billing_name or '',
            'ba':                  billing_add or '',
            'ba2':                 billing_add2 or '',
            'bc':                  billing_city or '',
            'bz':                  billing_zip or '',
            'billing_state':       billing_state or '',
            'billing_country':     billing_country or '',
            'bill_vat':            bil_vat or '',
            'lot':                 invoice.selected_dc_id.move_line_ids.lot_id if invoice.selected_dc_id else [],
            'delivery_dc_numbers': [dc.dc_number for dc in invoice.delivery_ids if dc.dc_number] if invoice.delivery_ids else ([invoice.selected_dc_id.dc_number] if invoice.selected_dc_id and invoice.selected_dc_id.dc_number else []),
            'delivery_count':      len([dc for dc in invoice.delivery_ids if dc.dc_number]) if invoice.delivery_ids else (1 if invoice.selected_dc_id and invoice.selected_dc_id.dc_number else 0),
        })

        pdf_bytes = HTML(string=html_content, base_url=f'file://{template_dir}/').write_pdf()
        return pdf_bytes


    # ============================================================
    # REPLACE generate_invoice_print_report() with this:
    # ============================================================

    def generate_invoice_print_report(self):
        """Generate Invoice Print Report - runs in background thread"""
        domain = [('move_type', '=', 'out_invoice')]
        if self.invoice_print_ids:
            domain.append(('id', 'in', self.invoice_print_ids.ids))
        if self.invoice_print_from_date:
            domain.append(('invoice_date', '>=', self.invoice_print_from_date))
        if self.invoice_print_to_date:
            domain.append(('invoice_date', '<=', self.invoice_print_to_date))
        if self.invoice_print_customer_id:
            domain.append(('partner_id', '=', self.invoice_print_customer_id.id))
        if self.invoice_print_status:
            domain.append(('state', '=', self.invoice_print_status))
        else:
            domain.append(('state', '=', 'posted'))

        invoices = self.env['account.move'].search(domain, order='invoice_date asc')
        if not invoices:
            raise ValidationError('No invoices found matching the selected filters.')

        # Build filter-based filename  soundharya
        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.invoice_print_ids:
            names = self.invoice_print_ids.mapped('name')
            parts.append(_sanitize('_'.join(names)) if len(names) <= 3 else f'{len(names)}_Invoices')
        if self.invoice_print_from_date:
            parts.append(self.invoice_print_from_date.strftime('%d%m%Y'))
        if self.invoice_print_to_date:
            parts.append('to_' + self.invoice_print_to_date.strftime('%d%m%Y'))
        if self.invoice_print_customer_id:
            parts.append(_sanitize(self.invoice_print_customer_id.name))
        if self.invoice_print_status:
            parts.append(self.invoice_print_status)
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        inv_filename = f'Invoice_Print_{"_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")}.{ext}'

        invoice_count = len(invoices)
        self.write({'pdf_file': False, 'pdf_filename': 'Generating...'})
        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        invoice_ids = invoices.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._invoice_print_background,
            args=(wizard_id, db, uid, invoice_ids, ctx, inv_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Generating {fmt_label} for {invoice_count} invoice(s) in the background.',
                'type': 'info',
                'sticky': True,
            }
        }

    def _invoice_print_background(self, wizard_id, db, uid, invoice_ids, ctx, inv_filename=None, download_type='pdf'):
        import odoo, io, zipfile, time

        def _write_result(file_b64, filename):
            # Fresh write cursor with retry — avoids serialization conflict from long-held read cursor
            for attempt in range(5):
                try:
                    with odoo.registry(db).cursor() as write_cr:
                        write_env = odoo.api.Environment(write_cr, uid, ctx)
                        write_env['reports.wizard'].browse(wizard_id).write({
                            'pdf_file': file_b64,
                            'pdf_filename': filename,
                        })
                        write_cr.commit()
                    return
                except Exception as e:
                    if attempt < 4 and ('serialize' in str(e).lower() or '40001' in str(e)):
                        time.sleep(0.3 * (2 ** attempt))
                        continue
                    _logger.error("Invoice print: write result failed after %d attempts: %s", attempt + 1, str(e), exc_info=True)
                    raise

        def _write_error(error_msg):
            for attempt in range(3):
                try:
                    with odoo.registry(db).cursor() as err_cr:
                        err_env = odoo.api.Environment(err_cr, uid, ctx)
                        err_env['reports.wizard'].browse(wizard_id).write({
                            'pdf_file': False,
                            'pdf_filename': f'ERROR: {str(error_msg)[:250]}',
                        })
                        err_cr.commit()
                    return
                except Exception:
                    if attempt < 2:
                        time.sleep(0.2 * (2 ** attempt))

        try:
            # Phase 1: read data + render PDF — cursor never writes to reports.wizard
            with odoo.registry(db).cursor() as read_cr:
                env = odoo.api.Environment(read_cr, uid, ctx)
                wizard = env['reports.wizard'].browse(wizard_id)
                invoices = env['account.move'].browse(invoice_ids)

                errors = []
                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for invoice in invoices:
                            try:
                                pdf_bytes = wizard._render_invoice_pdf_bytes(invoice)
                                if not pdf_bytes:
                                    errors.append(f'{invoice.name}: Empty PDF')
                                    continue
                                entry_name = f"{invoice.name.replace('/', '_')}.pdf"
                                zf.writestr(entry_name, pdf_bytes)
                                count += 1
                            except Exception as e:
                                errors.append(f'{invoice.name}: {str(e)[:200]}')
                                _logger.error("Invoice ZIP render failed for %s: %s", invoice.name, str(e), exc_info=True)
                    if not count:
                        err_detail = '; '.join(errors[:3]) if errors else 'Unknown error'
                        raise Exception(f'No invoice PDFs could be generated. {err_detail}')
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for invoice in invoices:
                        try:
                            pdf_bytes = wizard._render_invoice_pdf_bytes(invoice)
                            if not pdf_bytes:
                                errors.append(f'{invoice.name}: Empty PDF')
                                continue
                            reader = PdfReader(io.BytesIO(pdf_bytes))
                            for page in reader.pages:
                                writer.add_page(page)
                        except Exception as e:
                            errors.append(f'{invoice.name}: {str(e)[:200]}')
                            _logger.error("Invoice PDF render failed for %s: %s", invoice.name, str(e), exc_info=True)
                    if not writer.pages:
                        err_detail = '; '.join(errors[:3]) if errors else 'Unknown error'
                        raise Exception(f'No invoice PDFs could be generated. {err_detail}')
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()
            # read_cr closed here — no write conflict possible

            # Phase 2: write in a fresh cursor with retry
            filename = inv_filename or f'Invoice_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
            _write_result(base64.b64encode(output_bytes).decode('utf-8'), filename)

        except Exception as e:
            _logger.error("Invoice print background failed: %s", str(e), exc_info=True)
            _write_error(str(e))


    def generate_po_print_report(self):
        """Generate PO Print Report - runs in background thread"""
        domain = []
        if self.po_print_ids:
            domain.append(('id', 'in', self.po_print_ids.ids))
        if self.po_print_from_date:
            domain.append(('create_date_only', '>=', self.po_print_from_date))
        if self.po_print_to_date:
            domain.append(('create_date_only', '<=', self.po_print_to_date))
        if self.po_print_vendor_id:
            domain.append(('partner_id', '=', self.po_print_vendor_id.id))
        if self.po_print_status:
            domain.append(('state', '=', self.po_print_status))

        purchase_orders = self.env['purchase.order'].search(domain, order='create_date asc')
        if not purchase_orders:
            raise ValidationError('No Purchase Orders found matching the selected filters.')

        # Build filter-based filename  soundharya
        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.po_print_ids:
            names = self.po_print_ids.mapped('name')
            parts.append(_sanitize('_'.join(names)) if len(names) <= 3 else f'{len(names)}_POs')
        if self.po_print_from_date:
            parts.append(self.po_print_from_date.strftime('%d%m%Y'))
        if self.po_print_to_date:
            parts.append('to_' + self.po_print_to_date.strftime('%d%m%Y'))
        if self.po_print_vendor_id:
            parts.append(_sanitize(self.po_print_vendor_id.name))
        if self.po_print_status:
            parts.append(self.po_print_status)
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        po_filename = f'PO_Print_{"_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")}.{ext}'

        self.write({'pdf_file': False, 'pdf_filename': False})
        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        po_ids = purchase_orders.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._po_print_background,
            args=(wizard_id, db, uid, po_ids, ctx, po_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Your PO {fmt_label} is being generated in the background.',
                'type': 'info',
                'sticky': True,
            }
        }

    def _po_print_background(self, wizard_id, db, uid, po_ids, ctx, po_filename=None, download_type='pdf'):
        import odoo, io, zipfile, time

        def _write_result(file_b64, filename):
            # Fresh write cursor with retry — avoids serialization conflict from long-held read cursor
            for attempt in range(5):
                try:
                    with odoo.registry(db).cursor() as write_cr:
                        write_env = odoo.api.Environment(write_cr, uid, ctx)
                        write_env['reports.wizard'].browse(wizard_id).write({
                            'pdf_file': file_b64,
                            'pdf_filename': filename,
                        })
                        write_cr.commit()
                    return
                except Exception as e:
                    if attempt < 4 and ('serialize' in str(e).lower() or '40001' in str(e)):
                        time.sleep(0.3 * (2 ** attempt))
                        continue
                    _logger.error("PO print: write result failed after %d attempts: %s", attempt + 1, str(e), exc_info=True)
                    raise

        def _write_error(error_msg):
            for attempt in range(3):
                try:
                    with odoo.registry(db).cursor() as err_cr:
                        err_env = odoo.api.Environment(err_cr, uid, ctx)
                        err_env['reports.wizard'].browse(wizard_id).write({
                            'pdf_file': False,
                            'pdf_filename': f'ERROR: {str(error_msg)[:250]}',
                        })
                        err_cr.commit()
                    return
                except Exception:
                    if attempt < 2:
                        time.sleep(0.2 * (2 ** attempt))

        try:
            # Phase 1: read data + render PDF — cursor never writes to reports.wizard
            with odoo.registry(db).cursor() as read_cr:
                env = odoo.api.Environment(read_cr, uid, ctx)
                purchase_orders = env['purchase.order'].browse(po_ids)

                errors = []
                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for po in purchase_orders:
                            try:
                                pdf_bytes = po.generate_purchase_pdf(po.id)
                                if pdf_bytes:
                                    entry_name = f"{po.name.replace('/', '_')}.pdf"
                                    zf.writestr(entry_name, pdf_bytes)
                                    count += 1
                                else:
                                    errors.append(f'PO {po.name}: Empty PDF')
                            except Exception as e:
                                errors.append(f'PO {po.name}: {str(e)[:100]}')
                                _logger.error("PO PDF render failed for %s: %s", po.name, str(e))
                    if not count:
                        err_detail = '; '.join(errors[:3]) if errors else 'Unknown error'
                        raise Exception(f'No PO PDFs could be generated. Errors: {err_detail}')
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for po in purchase_orders:
                        try:
                            pdf_bytes = po.generate_purchase_pdf(po.id)
                            if pdf_bytes:
                                reader = PdfReader(io.BytesIO(pdf_bytes))
                                for page in reader.pages:
                                    writer.add_page(page)
                            else:
                                errors.append(f'PO {po.name}: Empty PDF')
                        except Exception as e:
                            errors.append(f'PO {po.name}: {str(e)[:100]}')
                            _logger.error("PO PDF render failed for %s: %s", po.name, str(e))
                    if not writer.pages:
                        err_detail = '; '.join(errors[:3]) if errors else 'Unknown error'
                        raise Exception(f'No PO PDFs could be generated. Errors: {err_detail}')
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()
            # read_cr closed here — no write conflict possible

            # Phase 2: write in a fresh cursor with retry
            filename = po_filename or f'PO_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
            _write_result(base64.b64encode(output_bytes).decode('utf-8'), filename)

        except Exception as e:
            _logger.error("PO print background failed: %s", str(e), exc_info=True)
            _write_error(str(e))

    # ============================================================
    # Quotation Print Report
    # ============================================================

    def _render_quotation_pdf_bytes(self, quotation):
        """
        Renders the Quotation PDF bytes directly by calling the same
        logic as the Print button on the quotation form.
        Replicates the controller/method that generates the quotation PDF.
        """
        import os
        from jinja2 import Environment, FileSystemLoader
        from weasyprint import HTML

        # ── Resolve template path ─────────────────────────────────
        # reports_wizard.py → hrms_dashboard/models/
        # Template is in   → accounting/static/templates/ or hrms_dashboard/static/templates/
        current_file  = os.path.abspath(__file__)
        models_dir    = os.path.dirname(current_file)
        hrms_root     = os.path.dirname(models_dir)
        custom_addons = os.path.dirname(hrms_root)

        # Try hrms_dashboard first, then accounting module
        template_name = 'static_quotation_template.html'
        template_dir  = os.path.join(custom_addons, 'zigma_erp', 'static', 'templates')

        html_path = os.path.join(template_dir, template_name)

        if not os.path.exists(html_path):
            raise ValidationError(
                f'Quotation template not found at: {html_path}\n'
                f'Please check your static/templates folder for the quotation HTML file.'
            )

        env_j2   = Environment(loader=FileSystemLoader(template_dir))
        template = env_j2.get_template(template_name)

        # ── Build context (same data the quotation print button uses) ─
        company = quotation.company_id or self.env.company

        # Tax calculations
        total_cgst = total_sgst = total_igst = 0.0
        is_igst = False
        for item in quotation.item_line_ids:
            total_cgst += getattr(item, 'cgst_amount', 0.0) or 0.0
            total_sgst += getattr(item, 'sgst_amount', 0.0) or 0.0
            total_igst += getattr(item, 'igst_amount', 0.0) or 0.0
            for tax in (item.tax_id if hasattr(item, 'tax_id') else []):
                if tax.tax_group_id.name == 'IGST':
                    is_igst = True
                    break
            if is_igst:
                break

        # Amount in words
        currency = quotation.currency_id or self.env.company.currency_id
        total_amount = quotation.amount_total or 0.0
        try:
            amount_in_words = currency.amount_to_text(total_amount)
            amount_in_words = amount_in_words.replace('-', ' ').replace(',', '') + ' Only'
        except Exception:
            amount_in_words = ''

        gst_percentage, gst_label, payment_terms_map = self._get_quotation_gst_info(quotation)
        payment_sel = getattr(quotation, 'payment_terms_selection', '') or ''
        payment_text = payment_terms_map.get(payment_sel, payment_sel)
        employee = quotation.employee_id

        validity_days = getattr(quotation, 'validity_days', '') or ''
        quote_validity = getattr(quotation, 'quote_validity', None)
        validity_terms = (
            f"This quote is valid for {validity_days} days from the date of the quote "
            f"(Valid until: {quote_validity.strftime('%d/%m/%Y') if quote_validity else ''})"
        )

        delivery_time = getattr(quotation, 'delivery_time', '') or ''
        amount_in_words = getattr(quotation, 'amount_in_words', '') or ''

        html_content = template.render({
            # Company Information
            'company_corporate_address': '7, New No. 2, Lakshmi Colony North Crescent Road, T. Nagar, Chennai - 600 017',
            'company_registered_address': '747, Amara Complex, S.K.C Road, Erode-638001',
            'company_gstin': company.vat or '33AACCT0611K1ZQ',
            'company_cin': company.company_registry or 'U72200TZ2004PTC011138',
            'company_pan': 'AACCT0611K',
            'company_email': 'info@zigmaindia.com',
            'company_website': 'www.zigmaindia.com',
            'company_phone': '9791933074/80/83,',

            # Quote Reference
            'quote_number': quotation.name or '',
            'quote_date': quotation.quotation_date.strftime('%d/%m/%Y') if quotation.quotation_date else '',
            'authorized_signatory_name': employee.name if employee else '',

            # Customer Information
            'customer_company': quotation.customer_id.name or '',
            'customer_address': '{}, {}, {} - {}'.format(
                quotation.customer_id.street or '',
                quotation.customer_id.street2 or '',
                quotation.customer_id.city or '',
                quotation.customer_id.zip or '',
            ),
            'customer_contact_person': quotation.contact_person.name if quotation.contact_person else '',
            'customer_mobile': (quotation.contact_person.phone or quotation.contact_person.mobile or '') if quotation.contact_person else '',
            'customer_email': (
                quotation.customer_id.email
                or (quotation.contact_person.email if quotation.contact_person else '')
                or ''
            ),

            # Products
            'products': self._build_quotation_products(quotation),

            # Pricing
            'subtotal': self._format_indian(quotation.amount_untaxed),
            'gst_percentage': gst_percentage,
            'gst_label': gst_label,
            'gst_amount': self._format_indian(quotation.amount_tax),
            'grand_total': self._format_indian(quotation.amount_total),
            'amount_in_words': amount_in_words,

            # Terms
            'validity_terms': validity_terms,
            'payment_terms': payment_text,
            'delivery_terms': 'Within {} weeks from the date of PO released'.format(delivery_time),
            'warranty_terms': 'As mentioned above',
            'po_details': 'PO should contain Quote number ({}), "complete bill to" & "ship to" address with contact details'.format(quotation.name or ''),
            'closing_note': 'We look forward to your valuable order. If you have any clarification, please feel free to call or mail us directly.',

            # Signature
            'authorized_signatory_name': employee.name if employee else '',
            'authorized_signatory_contact': (employee.mobile_phone or employee.work_phone or '') if employee else '',
            'authorized_signatory_email': (employee.work_email or '') if employee else '',
            'date_generated': datetime.now().strftime('%d/%m/%Y'),
        })

        pdf_bytes = HTML(string=html_content, base_url=f'file://{template_dir}/').write_pdf()
        return pdf_bytes

    def _format_indian(self, amount):
        """Format number in Indian numbering system (e.g. 1,32,000.00)."""
        try:
            s = "{:.2f}".format(float(amount))
            integer_part, decimal_part = s.split('.')
            negative = integer_part.startswith('-')
            if negative:
                integer_part = integer_part[1:]
            if len(integer_part) <= 3:
                result = integer_part
            else:
                last_three = integer_part[-3:]
                remaining = integer_part[:-3]
                groups = []
                while remaining:
                    groups.append(remaining[-2:] if len(remaining) >= 2 else remaining)
                    remaining = remaining[:-2]
                groups.reverse()
                result = ','.join(groups) + ',' + last_three
            return ('-' if negative else '') + result + '.' + decimal_part
        except Exception:
            return str(amount)

    def _build_quotation_products(self, quotation):
        """Build products list matching static_quotation_template.html format exactly."""
        products = []
        for item in quotation.item_line_ids:
            try:
                product_lines = quotation.product_line_ids.filtered(
                    lambda p: p.item_details_id == item.item_details_id
                )
                specifications = []
                for prod_line in product_lines:
                    if prod_line.product_id:
                        part = getattr(prod_line, 'part_number', '') or 'Not Available'
                        part_type = getattr(prod_line, 'part_number_type', '') or ''
                        detailed = getattr(prod_line.product_id, 'detailed', '') or prod_line.product_id.name or ''
                        specifications.append({
                            'part': part,
                            'part_number_type': part_type,
                            'description': detailed,
                        })

                item_name = getattr(item, 'item_name', '') or ''
                if not item_name and item.item_details_id:
                    item_name = getattr(item.item_details_id, 'item_name', '') or ''
                item_description = getattr(item, 'description', '') or ''
                if not item_description and item.item_details_id:
                    item_description = getattr(item.item_details_id, 'description', '') or ''

                qty = getattr(item, 'quotation_qty', 0) or 0
                price_subtotal = getattr(item, 'price_subtotal', 0.0) or 0.0
                unit_price = (price_subtotal / qty) if qty else price_subtotal
                period = getattr(item, 'period', '') or ''

                products.append({
                    'name': item_name,
                    'description': item_description,
                    'qty': qty,
                    'unit_price': self._format_indian(unit_price),
                    'total_price': self._format_indian(price_subtotal),
                    'warranty': '{} onsite Warranty'.format(period) if period else '',
                    'specifications': specifications,
                    'spec_title': 'Product {} - {}'.format(len(products) + 1, item_name),
                })
            except Exception as e:
                _logger.warning('Skipped quotation item due to error: %s', str(e))
        return products

    def _get_quotation_gst_info(self, quotation):
        """Return (gst_percentage, gst_label) from first tax found."""
        payment_terms_map = {
            'immediate': 'Immediate',
            'against_delivery': 'Against Delivery',
            '100_advance': '100% Advance Payment',
            '50_advance_50_delivery': '50% Advance Payment and 50% Payment Against Delivery and Installation',
            '100_delivery_installation': '100% Payment Against Delivery and Installation',
            '7_days_invoice': '7 Days from Date of Invoice',
            '15_days_invoice': '15 Days from Date of Invoice',
            '30_days_invoice': '30 Days from Date of Invoice',
        }
        gst_percentage = ''
        gst_label = 'GST'
        for line in quotation.item_line_ids:
            for tax in line.tax_id:
                gst_percentage = int(tax.amount) if tax.amount == int(tax.amount) else tax.amount
                combined = ((tax.name or '') + ' ' + (tax.tax_group_id.name or '')).upper()
                if 'IGST' in combined:
                    gst_label = 'IGST'
                elif 'CGST' in combined:
                    gst_label = 'CGST'
                elif 'SGST' in combined:
                    gst_label = 'SGST'
                break
            if gst_percentage:
                break
        return gst_percentage, gst_label, payment_terms_map

    def _quotation_print_background(self, wizard_id, db, uid, quotation_ids, ctx, quot_filename=None, download_type='pdf'):
        """Runs quotation PDF generation in a background thread."""
        import odoo, io, zipfile
        try:
            with odoo.registry(db).cursor() as new_cr:
                env = odoo.api.Environment(new_cr, uid, ctx)
                wizard = env['reports.wizard'].browse(wizard_id)
                quotations = env['quotation.management'].browse(quotation_ids)

                errors = []
                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for quotation in quotations:
                            try:
                                pdf_b = wizard._render_quotation_pdf_bytes(quotation)
                                entry_name = f"{quotation.name.replace('/', '_')}.pdf"
                                zf.writestr(entry_name, pdf_b)
                                count += 1
                            except Exception as e:
                                errors.append(f'{quotation.name}: {str(e)}')
                    if not count:
                        raise Exception('Could not generate any Quotation PDFs.\n' + '\n'.join(errors))
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for quotation in quotations:
                        try:
                            pdf_b = wizard._render_quotation_pdf_bytes(quotation)
                            reader = PdfReader(io.BytesIO(pdf_b))
                            for page in reader.pages:
                                writer.add_page(page)
                        except Exception as e:
                            errors.append(f'{quotation.name}: {str(e)}')
                    if not writer.pages:
                        raise Exception('Could not generate any Quotation PDFs.\n' + '\n'.join(errors))
                    if errors:
                        _logger.warning('Quotation Print merge errors: %s', '; '.join(errors))
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()

                filename = quot_filename or f'Quotation_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
                wizard.write({
                    'pdf_file': base64.b64encode(output_bytes).decode('utf-8'),
                    'pdf_filename': filename,
                })
                new_cr.commit()

        except Exception as e:
            _logger.error("Quotation print background failed: %s", str(e), exc_info=True)

    def generate_quotation_print_report(self):
        """Generate Quotation Print Report in background thread (same pattern as other print reports)."""
        printable_states = ['approved', 'quote_sent', 'confirm']
        domain = []

        if self.quotation_print_ids:
            domain.append(('id', 'in', self.quotation_print_ids.ids))
        if self.quotation_print_from_date:
            domain.append(('quotation_date', '>=', self.quotation_print_from_date))
        if self.quotation_print_to_date:
            domain.append(('quotation_date', '<=', self.quotation_print_to_date))
        if self.quotation_print_customer_id:
            domain.append(('customer_id', '=', self.quotation_print_customer_id.id))
        if self.quotation_print_status:
            domain.append(('state', '=', self.quotation_print_status))
        else:
            domain.append(('state', 'in', printable_states))

        quotations = self.env['quotation.management'].search(domain)
        if not quotations:
            raise ValidationError(
                'No Quotations found in Approved / Quotation Sent / Order Confirmed state '
                'matching the selected filters.'
            )

        # Build filter-based filename  soundharya
        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.quotation_print_ids:
            names = self.quotation_print_ids.mapped('name')
            parts.append(_sanitize('_'.join(names)) if len(names) <= 3 else f'{len(names)}_Quotations')
        if self.quotation_print_from_date:
            parts.append(self.quotation_print_from_date.strftime('%d%m%Y'))
        if self.quotation_print_to_date:
            parts.append('to_' + self.quotation_print_to_date.strftime('%d%m%Y'))
        if self.quotation_print_customer_id:
            parts.append(_sanitize(self.quotation_print_customer_id.name))
        if self.quotation_print_status:
            parts.append(self.quotation_print_status)
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        quot_filename = f'Quotation_Print_{"_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")}.{ext}'

        self.write({'pdf_file': False, 'pdf_filename': False})

        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        quotation_ids = quotations.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._quotation_print_background,
            args=(wizard_id, db, uid, quotation_ids, ctx, quot_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Your Quotation {fmt_label} is being generated in the background.',
                'type': 'info',
                'sticky': True,
            }
        }


    # def generate_quotation_print_report(self):
    #     """
    #     Generate Quotation Print Report.
    #     Only shows quotations in states: approved, quote_sent, confirm.
    #     Single quotation → calls the quotation's print method directly.
    #     Multiple quotations → renders each PDF and merges into one.
    #     """
    #     import io
    #     try:
    #         from pypdf import PdfWriter, PdfReader
    #     except ImportError:
    #         from PyPDF2 import PdfWriter, PdfReader

    #     # ── Default: only show printable states ──────────────────────
    #     printable_states = ['approved', 'quote_sent', 'confirm']

    #     domain = []

    #     # if self.quotation_print_id:
    #     #     domain.append(('id', '=', self.quotation_print_id.id))

    #     if self.quotation_print_ids:
    #         domain.append(('id', 'in', self.quotation_print_ids.ids))
    #     if self.quotation_print_from_date:
    #         domain.append(('quotation_date', '>=', self.quotation_print_from_date))
    #     if self.quotation_print_to_date:
    #         domain.append(('quotation_date', '<=', self.quotation_print_to_date))
    #     if self.quotation_print_customer_id:
    #         domain.append(('customer_id', '=', self.quotation_print_customer_id.id))

    #     # Status filter: if specific status chosen use it, else default to all printable states
    #     if self.quotation_print_status:
    #         domain.append(('state', '=', self.quotation_print_status))
    #     else:
    #         domain.append(('state', 'in', printable_states))

    #     quotations = self.env['quotation.management'].search(domain, order='quotation_date asc')

    #     if not quotations:
    #         raise ValidationError(
    #             'No Quotations found in Approved / Quotation Sent / Order Confirmed state '
    #             'matching the selected filters.'
    #         )

    #     # ── Single quotation: try the model's own print method ───────
    #     if len(quotations) == 1:
    #         quotation = quotations[0]
    #         # Try the model's print method (same as the Print button)
    #         for method_name in ['action_print_quotation', 'generate_quotation_pdf',
    #                             'action_print_pdf', 'print_quotation']:
    #             if hasattr(quotation, method_name):
    #                 try:
    #                     return getattr(quotation, method_name)()
    #                 except Exception:
    #                     pass

    #         # Fallback: render directly via weasyprint
    #         try:
    #             pdf_bytes = self._render_quotation_pdf_bytes(quotation)
    #             filename = f'Quotation_{quotation.name.replace("/", "_")}.pdf'
    #             attachment = self.env['ir.attachment'].create({
    #                 'name': filename,
    #                 'type': 'binary',
    #                 'datas': base64.b64encode(pdf_bytes),
    #                 'res_model': 'reports.wizard',
    #                 'res_id': self.id,
    #                 'mimetype': 'application/pdf',
    #             })
    #             return {
    #                 'type': 'ir.actions.act_url',
    #                 'url': f'/web/content/{attachment.id}?download=true',
    #                 'target': 'new',
    #             }
    #         except Exception as e:
    #             raise ValidationError(f'Could not generate Quotation PDF: {str(e)}')

    #     # ── Multiple quotations: render each and merge ────────────────
    #     writer = PdfWriter()
    #     errors = []

    #     for quotation in quotations:
    #         try:
    #             pdf_bytes = self._render_quotation_pdf_bytes(quotation)
    #             if not pdf_bytes:
    #                 errors.append(f'Quotation {quotation.name}: Empty PDF')
    #                 continue
    #             reader = PdfReader(io.BytesIO(pdf_bytes))
    #             for page in reader.pages:
    #                 writer.add_page(page)
    #         except Exception as e:
    #             errors.append(f'Quotation {quotation.name}: {str(e)}')
    #             continue

    #     if not writer.pages:
    #         raise ValidationError(
    #             'Could not generate any Quotation PDFs.\n\nErrors:\n' + '\n'.join(errors)
    #         )

    #     output_buffer = io.BytesIO()
    #     writer.write(output_buffer)
    #     output_buffer.seek(0)
    #     merged_pdf_bytes = output_buffer.read()

    #     filename = f'Quotation_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
    #     attachment = self.env['ir.attachment'].create({
    #         'name': filename,
    #         'type': 'binary',
    #         'datas': base64.b64encode(merged_pdf_bytes),
    #         'res_model': 'reports.wizard',
    #         'res_id': self.id,
    #         'mimetype': 'application/pdf',
    #     })

    #     if errors:
    #         import logging
    #         _logger = logging.getLogger(__name__)
    #         _logger.warning('Quotation Print merge: some PDFs failed: %s', '; '.join(errors))

    #     return {
    #         'type': 'ir.actions.act_url',
    #         'url': f'/web/content/{attachment.id}?download=true',
    #         'target': 'new',
    #     }


    # def generate_installation_print_report(self):
    #     """Generate Installation Print Report - prints all matching Installation task PDFs merged into one"""
    #     import io
    #     try:
    #         from pypdf import PdfWriter, PdfReader
    #     except ImportError:
    #         from PyPDF2 import PdfWriter, PdfReader

    #     # Build domain - only installation projects
    #     installation_projects = self.env['project.project'].search([
    #         ('name', 'ilike', 'INS')
    #     ])

    #     domain = [('project_id', 'in', installation_projects.ids)]

    #     if self.installation_print_from_date:
    #         domain.append(('create_date', '>=', self.installation_print_from_date))

    #     if self.installation_print_to_date:
    #         domain.append(('create_date', '<=', self.installation_print_to_date))

    #     if self.installation_print_customer_id:
    #         domain.append(('project_id.partner_id', '=', self.installation_print_customer_id.id))

    #     if self.installation_print_project_id:
    #         domain.append(('project_id', '=', self.installation_print_project_id.id))

    #     if self.installation_print_task_id:
    #         domain.append(('id', '=', self.installation_print_task_id.id))

    #     if self.installation_print_status:
    #         domain.append(('state', '=', self.installation_print_status))

    def generate_installation_print_report(self):
        """Generate Installation Print Report - runs in background thread"""
        installation_projects = self.env['project.project'].search([('name', 'ilike', 'INS')])
        domain = [('project_id', 'in', installation_projects.ids)]

        if self.installation_print_from_date:
            domain.append(('create_date', '>=', self.installation_print_from_date))
        if self.installation_print_to_date:
            domain.append(('create_date', '<=', self.installation_print_to_date))
        if self.installation_print_customer_id:
            domain.append(('project_id.partner_id', '=', self.installation_print_customer_id.id))
        if self.installation_print_project_ids:
            domain.append(('project_id', 'in', self.installation_print_project_ids.ids))
        if self.installation_print_task_ids:
            domain.append(('id', 'in', self.installation_print_task_ids.ids))
        if self.installation_print_stage_id:
            domain.append(('stage_id', '=', self.installation_print_stage_id.id))
        if self.installation_print_state:
            domain.append(('state', '=', self.installation_print_state))

        tasks = self.env['project.task'].search(domain, order='create_date asc')
        if not tasks:
            raise ValidationError('No Installation tasks found matching the selected filters.')

        # Build filter-based filename  soundharya
        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.installation_print_from_date:
            parts.append(self.installation_print_from_date.strftime('%d%m%Y'))
        if self.installation_print_to_date:
            parts.append('to_' + self.installation_print_to_date.strftime('%d%m%Y'))
        if self.installation_print_customer_id:
            parts.append(_sanitize(self.installation_print_customer_id.name))
        if self.installation_print_project_ids:
            names = self.installation_print_project_ids.mapped('name')
            parts.append(_sanitize('_'.join(names)) if len(names) <= 2 else f'{len(names)}_Projects')
        if self.installation_print_task_ids:
            parts.append(f'{len(self.installation_print_task_ids)}_Tasks')
        if self.installation_print_stage_id:
            parts.append(_sanitize(self.installation_print_stage_id.name, 15))
        if self.installation_print_state:
            parts.append(self.installation_print_state)
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        inst_filename = f'Installation_Print_{"_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")}.{ext}'

        self.write({'pdf_file': False, 'pdf_filename': False})
        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        task_ids = tasks.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._installation_print_background,
            args=(wizard_id, db, uid, task_ids, ctx, inst_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Your Installation {fmt_label} is being generated in the background.',
                'type': 'info',
                'sticky': True,
            }
        }

    def _installation_print_background(self, wizard_id, db, uid, task_ids, ctx, inst_filename=None, download_type='pdf'):
        import odoo, io, zipfile
        try:
            with odoo.registry(db).cursor() as new_cr:
                env = odoo.api.Environment(new_cr, uid, ctx)
                wizard = env['reports.wizard'].browse(wizard_id)
                tasks = env['project.task'].browse(task_ids)

                errors = []
                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for task in tasks:
                            try:
                                pdf_bytes = wizard._render_installation_pdf_bytes(task)
                                if not pdf_bytes:
                                    errors.append(f'Task {task.name}: Empty PDF')
                                    continue
                                entry_name = f"{task.name.replace('/', '_')}.pdf"
                                zf.writestr(entry_name, pdf_bytes)
                                count += 1
                            except Exception as e:
                                errors.append(f'Task {task.name}: {str(e)}')
                    if not count:
                        raise Exception('No Installation PDFs could be generated.')
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for task in tasks:
                        try:
                            pdf_bytes = wizard._render_installation_pdf_bytes(task)
                            if not pdf_bytes:
                                errors.append(f'Task {task.name}: Empty PDF')
                                continue
                            reader = PdfReader(io.BytesIO(pdf_bytes))
                            for page in reader.pages:
                                writer.add_page(page)
                        except Exception as e:
                            errors.append(f'Task {task.name}: {str(e)}')
                    if not writer.pages:
                        raise Exception('No Installation PDFs could be generated.')
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()

                filename = inst_filename or f'Installation_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
                wizard.write({
                    'pdf_file': base64.b64encode(output_bytes).decode('utf-8'),
                    'pdf_filename': filename,
                })
                new_cr.commit()
        except Exception as e:
            _logger.error("Installation print background failed: %s", str(e))


    def _render_installation_pdf_bytes(self, task):
        """
        Renders the Installation PDF bytes directly by replicating
        the logic in ZigmaPrintController.print_zigma_invoice()
        """
        import os
        from jinja2 import Environment, FileSystemLoader
        from weasyprint import HTML

        current_file  = os.path.abspath(__file__)
        models_dir    = os.path.dirname(current_file)      # hrms_dashboard/models/
        hrms_root     = os.path.dirname(models_dir)        # hrms_dashboard/
        custom_addons = os.path.dirname(hrms_root)         # custom-addons/
        odoo17_root   = os.path.dirname(custom_addons)     # /opt/odoo17/

        html_path = os.path.join(odoo17_root, 'addons', 'project', 'static', 'templates', 'zigma.html')

        if not os.path.exists(html_path):
            raise ValidationError(
                f'Installation template not found at: {html_path}'
            )

        env_j2 = Environment(loader=FileSystemLoader(os.path.dirname(html_path)))
        template = env_j2.get_template(os.path.basename(html_path))

        # Build product lines the same way as ZigmaPrintController._build_product_lines:
        # task.item_delivery_ids holds the description; task.product_delivery_ids holds
        # the stock moves (serial numbers). Passing product_delivery_ids directly gives
        # stock.move records that have no 'description' attribute → blank item column.
        move_map = {}
        for move in (task.product_delivery_ids or []):
            if move.item_details_id:
                move_map.setdefault(move.item_details_id.id, []).append(move)

        product_lines = []
        for line in (task.item_delivery_ids or []):
            moves = move_map.get(line.item_details_id.id, []) if line.item_details_id else []
            product = moves[0].product_id if moves else None
            serial_parts = [
                getattr(m, 'serial_numbers_display', '') or ''
                for m in moves
                if m.product_id and m.product_id.tracking == 'serial'
            ]
            serial_display = ', '.join(s for s in serial_parts if s)
            product_lines.append({
                'description': line.description or '',
                'product_id': product,
                'serial_numbers_display': serial_display,
                'product_uom_qty': line.product_uom_qty,
            })

        html = template.render({
            'in_no': task.name or '',
            'contact_company': task.installation_address or '',
            'gstin': task.partner_id.vat if task.partner_id and task.partner_id.vat else 'UNREGISTERED',
            'po_no': task.project_id.customer_po.client_order_ref if task.project_id.customer_po else '',
            'po_date': task.project_id.customer_po.po_issue_date.strftime('%d-%m-%Y') if task.project_id.customer_po else '',
            'dc_no': task.dc_id.dc_id.dc_number if task.dc_id else '',
            'dc_date': task.dc_id.dc_id.date_done.strftime('%d-%m-%Y') if task.dc_id else '',
            # 'product_lines': task.product_delivery_ids or [],
            'contact_person_name': task.dc_id.dc_id.consignee_separation_id.contact_person_name if task.dc_id and task.dc_id.dc_id and task.dc_id.dc_id.consignee_separation_id else '',
            'contact_person_phone': (task.dc_id.dc_id.consignee_separation_id.contact_person.phone or task.dc_id.dc_id.consignee_separation_id.contact_person.mobile or '') if task.dc_id and task.dc_id.dc_id and task.dc_id.dc_id.consignee_separation_id and task.dc_id.dc_id.consignee_separation_id.contact_person else '',
            'product_lines': product_lines,
        })

        inst_template_dir = os.path.dirname(html_path)
        pdf_bytes = HTML(string=html, base_url=f'file://{inst_template_dir}/').write_pdf()
        return pdf_bytes

    def generate_elcot_print_report(self):
        """Generate ELCOT Print Report - runs in background thread"""
        installation_projects = self.env['project.project'].search([('name', 'ilike', 'INS')])
        domain = [('project_id', 'in', installation_projects.ids)]

        if self.elcot_print_from_date:
            domain.append(('create_date', '>=', self.elcot_print_from_date))
        if self.elcot_print_to_date:
            domain.append(('create_date', '<=', self.elcot_print_to_date))
        if self.elcot_print_customer_id:
            domain.append(('project_id.partner_id', '=', self.elcot_print_customer_id.id))
        if self.elcot_print_project_ids:
            domain.append(('project_id', 'in', self.elcot_print_project_ids.ids))
        if self.elcot_print_task_ids:
            domain.append(('id', 'in', self.elcot_print_task_ids.ids))
        if self.elcot_print_state:
            domain.append(('state', '=', self.elcot_print_state))

        tasks = self.env['project.task'].search(domain, order='create_date asc')
        if not tasks:
            raise ValidationError('No Installation tasks found matching the selected filters.')

        def _sanitize(val, maxlen=20):
            import re
            return re.sub(r'[^A-Za-z0-9_-]', '_', str(val))[:maxlen].strip('_')

        parts = []
        if self.elcot_print_from_date:
            parts.append(self.elcot_print_from_date.strftime('%d%m%Y'))
        if self.elcot_print_to_date:
            parts.append('to_' + self.elcot_print_to_date.strftime('%d%m%Y'))
        if self.elcot_print_customer_id:
            parts.append(_sanitize(self.elcot_print_customer_id.name))
        if self.elcot_print_project_ids:
            names = self.elcot_print_project_ids.mapped('name')
            parts.append(_sanitize('_'.join(names)) if len(names) <= 2 else f'{len(names)}_Projects')
        if self.elcot_print_task_ids:
            parts.append(f'{len(self.elcot_print_task_ids)}_Tasks')
        download_type = self.print_download_type or 'pdf'
        ext = 'zip' if download_type == 'zip' else 'pdf'
        elcot_filename = f'ELCOT_Print_{"_".join(parts) if parts else datetime.now().strftime("%Y%m%d_%H%M%S")}.{ext}'

        self.write({'pdf_file': False, 'pdf_filename': False})
        wizard_id = self.id
        db = self.env.cr.dbname
        uid = self.env.uid
        task_ids = tasks.ids
        ctx = dict(self.env.context)

        t = threading.Thread(
            target=self._elcot_print_background,
            args=(wizard_id, db, uid, task_ids, ctx, elcot_filename, download_type),
            daemon=True
        )
        t.start()

        fmt_label = 'ZIP' if download_type == 'zip' else 'PDF'
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': f'Generating {fmt_label}',
                'message': f'Your ELCOT {fmt_label} is being generated in the background.',
                'type': 'info',
                'sticky': True,
            }
        }

    def _elcot_print_background(self, wizard_id, db, uid, task_ids, ctx, elcot_filename=None, download_type='pdf'):
        import odoo, io, zipfile
        try:
            with odoo.registry(db).cursor() as new_cr:
                env = odoo.api.Environment(new_cr, uid, ctx)
                wizard = env['reports.wizard'].browse(wizard_id)
                tasks = env['project.task'].browse(task_ids)

                errors = []
                if download_type == 'zip':
                    zip_buffer = io.BytesIO()
                    count = 0
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for task in tasks:
                            try:
                                pdf_bytes = wizard._render_elcot_pdf_bytes(task)
                                if not pdf_bytes:
                                    errors.append(f'Task {task.name}: Empty PDF')
                                    continue
                                entry_name = f"{task.name.replace('/', '_')}.pdf"
                                zf.writestr(entry_name, pdf_bytes)
                                count += 1
                            except Exception as e:
                                errors.append(f'Task {task.name}: {str(e)}')
                    if not count:
                        raise Exception('No ELCOT PDFs could be generated.')
                    zip_buffer.seek(0)
                    output_bytes = zip_buffer.read()
                else:
                    try:
                        from pypdf import PdfWriter, PdfReader
                    except ImportError:
                        from PyPDF2 import PdfWriter, PdfReader
                    writer = PdfWriter()
                    for task in tasks:
                        try:
                            pdf_bytes = wizard._render_elcot_pdf_bytes(task)
                            if not pdf_bytes:
                                errors.append(f'Task {task.name}: Empty PDF')
                                continue
                            reader = PdfReader(io.BytesIO(pdf_bytes))
                            for page in reader.pages:
                                writer.add_page(page)
                        except Exception as e:
                            errors.append(f'Task {task.name}: {str(e)}')
                    if not writer.pages:
                        raise Exception('No ELCOT PDFs could be generated.')
                    output_buffer = io.BytesIO()
                    writer.write(output_buffer)
                    output_buffer.seek(0)
                    output_bytes = output_buffer.read()

                filename = elcot_filename or f'ELCOT_Print_{datetime.now().strftime("%Y%m%d_%H%M%S")}.{"zip" if download_type == "zip" else "pdf"}'
                wizard.write({
                    'pdf_file': base64.b64encode(output_bytes).decode('utf-8'),
                    'pdf_filename': filename,
                })
                new_cr.commit()
        except Exception as e:
            _logger.error("ELCOT print background failed: %s", str(e))

    def _render_elcot_pdf_bytes(self, task):
        """Render the ELCOT Installation Certificate PDF bytes using the elcot.html template"""
        import os
        from jinja2 import Environment, FileSystemLoader
        from weasyprint import HTML

        current_file = os.path.abspath(__file__)
        models_dir = os.path.dirname(current_file)       # hrms_dashboard/models/
        hrms_root = os.path.dirname(models_dir)          # hrms_dashboard/
        custom_addons = os.path.dirname(hrms_root)       # custom-addons/

        template_dir = os.path.join(custom_addons, 'accounting', 'static', 'templates')
        html_path = os.path.join(template_dir, 'elcot.html')

        if not os.path.exists(html_path):
            raise ValidationError(f'ELCOT template not found at: {html_path}')

        env_j2 = Environment(loader=FileSystemLoader(template_dir))
        template = env_j2.get_template('elcot.html')

        # Billing Address
        billing_name = ''
        billing_address = ''
        partner = task.dc_id.sale_order_id.partner_id if task.dc_id and task.dc_id.sale_order_id else task.partner_id
        if partner:
            billing_name = partner.name or ''
            city_zip = ' '.join(filter(None, [partner.city or '', partner.zip or '']))
            billing_address = ', '.join(filter(None, [
                partner.street, partner.street2, city_zip,
                partner.state_id.name if partner.state_id else '',
            ]))

        # Consignee Address
        consignee_name = ''
        consignee_address = ''
        if task.dc_id and task.dc_id.dc_id and task.dc_id.dc_id.consignee_separation_id:
            sep = task.dc_id.dc_id.consignee_separation_id
            consignee_name = sep.customer_id.name or ''
            addr_partner = sep.location_id if sep.branch_applicable else sep.customer_id
            if addr_partner:
                city_zip = ' '.join(filter(None, [addr_partner.city or '', addr_partner.zip or '']))
                consignee_address = ', '.join(filter(None, [
                    addr_partner.street, addr_partner.street2, city_zip,
                    addr_partner.state_id.name if addr_partner.state_id else '',
                ]))
        else:
            consignee_address = task.installation_address or ''

        # Contact Person & No
        contact_person = ''
        contact_no = ''
        if task.dc_id and task.dc_id.sale_order_id and task.dc_id.sale_order_id.contact_person:
            cp = task.dc_id.sale_order_id.contact_person
            contact_person = cp.name or ''
            contact_no = cp.phone or cp.mobile or ''
        elif task.dc_id and task.dc_id.dc_id:
            consignee = task.dc_id.dc_id.partner_id
            if consignee.type == 'delivery':
                consignee_contact = consignee
            else:
                consignee_contact = consignee.child_ids.filtered(lambda c: c.type == 'contact')[:1]
            if consignee_contact:
                contact_person = consignee_contact.name or ''
                contact_no = consignee_contact.mobile or consignee_contact.phone or consignee.mobile or consignee.phone or ''

        # PO info
        po_no = ''
        po_date = ''
        if task.dc_id and task.dc_id.sale_order_id:
            sale_order = task.dc_id.sale_order_id
            po_no = sale_order.client_order_ref or ''
            if sale_order.po_issue_date:
                po_date = sale_order.po_issue_date.strftime('%d/%m/%Y')

        # DC info
        dc_no = ''
        dc_date = ''
        picking = task.dc_id.dc_id if task.dc_id else None
        if picking:
            dc_no = picking.dc_number or ''
            if picking.date_done:
                dc_date = picking.date_done.strftime('%d/%m/%Y')

        # Items
        items = []
        if picking and picking.item_line_ids:
            for item_line in picking.item_line_ids:
                moves = picking.move_ids_without_package.filtered(
                    lambda m: m.item_details_id == item_line.item_details_id
                )
                serial_numbers = []
                for move in moves:
                    if move.product_id.tracking == 'serial':
                        for ml in move.move_line_ids:
                            if ml.lot_id:
                                serial_numbers.append(ml.lot_id.name)
                items.append({
                    'description': item_line.description or item_line.item_name or '',
                    'hsn': ', '.join(serial_numbers),
                    'qty': item_line.order_qty or 0,
                    'rate': '',
                })

        html_content = template.render({
            'in_no': task.name or '',
            'is_igst': False,
            'billing_name': billing_name,
            'billing_address': billing_address,
            'consignee_name': consignee_name,
            'consignee_address': consignee_address,
            'contact_person': contact_person,
            'contact_no': contact_no,
            'po_no': po_no,
            'po_date': po_date,
            'dc_no': dc_no,
            'dc_date': dc_date,
            'items': items,
        })

        pdf_bytes = HTML(string=html_content, base_url=f'file://{template_dir}/').write_pdf()
        return pdf_bytes

    def _get_customer_name(self, enquiry, sale):
        """Helper to extract customer name from sale or enquiry"""
        if sale and sale.partner_id:
            return sale.partner_id.name
        if enquiry:
            if hasattr(enquiry, 'company_id') and enquiry.company_id:
                return enquiry.company_id.name
            if hasattr(enquiry, 'company') and enquiry.company:
                return enquiry.company
        return ''

    # Copy your existing helper methods
    def _write_lead_row(self, worksheet, row, enquiry, lead, cell_format, date_format):
        """Write a row with enquiry and lead data"""
        data = [
            enquiry.name or '',
            enquiry.create_date.date() if enquiry.create_date else '',
            lead.name or '',
            lead.employee_id.name if lead.employee_id else '',
            lead.partner_id.name if lead.partner_id else '',
            lead.company_partner.name if lead.company_partner else '',
            lead.date_order.date() if lead.date_order else '',
            '', '', '', '', '', '', '', ''  # Empty cells for quotation, sale order, DC data
        ]
        
        for col, value in enumerate(data):
            if col in [1, 6]:  # Date columns
                worksheet.write(row, col, value, date_format)
            else:
                worksheet.write(row, col, value, cell_format)
    
    def _write_quotation_row(self, worksheet, row, enquiry, lead, quotation, cell_format, date_format):
        """Write a row with enquiry, lead, and quotation data"""
        data = [
            enquiry.name or '',
            enquiry.create_date.date() if enquiry.create_date else '',
            lead.name or '',
            lead.employee_id.name if lead.employee_id else '',
            lead.partner_id.name if lead.partner_id else '',
            lead.company_partner.name if lead.company_partner else '',
            lead.date_order.date() if lead.date_order else '',
            quotation.name or '',
            quotation.quotation_date if quotation.quotation_date else '',
            '', '', '', '', '', '', ''  # Empty cells for sale order and DC data
        ]
        
        for col, value in enumerate(data):
            if col in [1, 6, 8]:  # Date columns
                worksheet.write(row, col, value, date_format)
            else:
                worksheet.write(row, col, value, cell_format)
    
    def _write_complete_row(self, worksheet, row, enquiry, lead, quotation, sale_order, cell_format, date_format):
        """Write a complete row with all data including DC and installation status"""
        # Get DC data for this sale order - Updated logic
        dcs = self.env['stock.picking'].search([
            ('sale_id', '=', sale_order.id),
            ('picking_type_id.code', '=', 'outgoing'),
            ('state', '!=', 'cancel')
        ])
        
        # If no DCs found with sale_id, try with origin field as fallback
        if not dcs:
            dcs = self.env['stock.picking'].search([
                ('origin', '=', sale_order.name),
                ('picking_type_id.code', '=', 'outgoing'),
                ('state', '!=', 'cancel')
            ])
        
        # Extract DC numbers and dates
        dc_numbers = []
        dc_dates = []
        
        for dc in dcs:
            # Use dc_number if available, otherwise use name
            dc_number = getattr(dc, 'dc_number', None) or dc.name
            dc_numbers.append(dc_number)
            
            # Use scheduled_date or date_done
            dc_date = dc.scheduled_date or dc.date_done
            if dc_date:
                dc_dates.append(dc_date.strftime('%d/%m/%Y'))
            else:
                dc_dates.append('')
        
        # Join the lists with commas
        dc_numbers_str = ', '.join(dc_numbers) if dc_numbers else ''
        dc_dates_str = ', '.join([date for date in dc_dates if date]) if dc_dates else ''
        
        # Get installation status from project updates
        installation_status = ''
        completion_date = ''
        
        # Method 1: Find project using customer_po field that matches sale order name
        projects = self.env['project.project'].search([
            ('customer_po', '=', sale_order.name)
        ])
        
        # Method 2: If no project found, try with customer_po containing sale order name
        if not projects:
            projects = self.env['project.project'].search([
                ('customer_po', 'ilike', sale_order.name)
            ])
        
        # Method 3: Try with sale_order_id if that field exists in project.project
        if not projects:
            projects = self.env['project.project'].search([
                ('sale_order_id', '=', sale_order.id)
            ])
        
        if projects:
            # Get the latest project update for each project
            for project in projects:
                updates = self.env['project.update'].search([
                    ('project_id', '=', project.id)
                ], order='date asc', limit=1)
                
                if updates:
                    # Get the status - convert selection value to display text
                    status_selection = dict(updates._fields['status'].selection)
                    installation_status = status_selection.get(updates.status, updates.status or '')
                    
                    # Get the date
                    completion_date = updates.date.strftime('%d/%m/%Y') if updates.date else ''
                    break  # Take the first project's update found
        
        # Prepare data row
        data = [
            enquiry.name or '',
            enquiry.create_date.date() if enquiry.create_date else '',
            lead.name or '',
            lead.employee_id.name if lead.employee_id else '',
            lead.partner_id.name if lead.partner_id else '',
            lead.company_partner.name if lead.company_partner else '',
            lead.date_order.date() if lead.date_order else '',
            quotation.name or '',
            quotation.quotation_date if quotation.quotation_date else '',
            sale_order.name or '',
            sale_order.date_order.date() if sale_order.date_order else '',
            dc_numbers_str,
            dc_dates_str,
            installation_status,
            completion_date
        ]
        
        # Write data to Excel
        for col, value in enumerate(data):
            if col in [1, 6, 8, 10]:  # Date columns (excluding completion_date as it's already formatted)
                worksheet.write(row, col, value, date_format)
            else:
                worksheet.write(row, col, value, cell_format)

    # ============================================================================
    # PRODUCT SUMMARY REPORT
    # ============================================================================

    def generate_product_summary_report(self):
        """Generate Product Summary Report Excel file"""
        domain = [('active', '=', True)]

        if self.product_summary_name:
            domain.append(('id', '=', self.product_summary_name.id))
        if self.product_summary_category_id:
            domain.append(('categ_id', 'child_of', self.product_summary_category_id.id))
        if self.product_summary_product_type:
            domain.append(('type', '=', self.product_summary_product_type))

        products = self.env['product.template'].search(domain, order='name asc')

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Product Summary')

        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
        })
        cell_format = workbook.add_format({
            'border': 1,
            'text_wrap': True,
            'valign': 'vcenter',
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1,
            'valign': 'vcenter',
        })
        center_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'align': 'center',
        })

        headers = [
            'Product Name', 'Internal Reference', 'Category', 'Product Type',
            'Part Number Type', 'Part Number', 'Cost Price', 'HSN Code', 'Qty On Hand',
        ]
        col_widths = [35, 20, 25, 18, 18, 25, 14, 15, 14]

        for idx, width in enumerate(col_widths):
            worksheet.set_column(idx, idx, width)
        worksheet.set_row(0, 25)

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        type_labels = {
            'consu': 'Consumable',
            'service': 'Service',
            'product': 'Storable Product',
        }

        preview_data = []
        row = 1
        for product in products:
            qty_on_hand = sum(
                product.product_variant_ids.mapped('qty_available')
            )

            hsn_code = ''
            if hasattr(product, 'l10n_in_hsn_code'):
                hsn_code = product.l10n_in_hsn_code or ''

            main_vendor = ''
            if product.seller_ids:
                main_vendor = product.seller_ids[0].partner_id.name or ''

            taxes = ', '.join(product.taxes_id.mapped('name')) if product.taxes_id else ''

            part_number_type_labels = {
                'auto_generate': 'Auto Generate',
                'manual': 'Manual',
            }
            raw_pnt = getattr(product, 'product_id_part_number_type', '') or ''
            part_number_type = part_number_type_labels.get(raw_pnt, raw_pnt)
            part_number = getattr(product, 'part_number', '') or ''

            row_data = [
                product.name or '',
                product.default_code or '',
                product.categ_id.complete_name if product.categ_id else '',
                type_labels.get(product.type, product.type or ''),
                part_number_type,
                part_number,
                product.standard_price,
                hsn_code,
                qty_on_hand,
            ]
    
            for col, value in enumerate(row_data):
                if col in [6, 8]:
                    worksheet.write(row, col, value, number_format)
                else:
                    worksheet.write(row, col, value, cell_format)

            preview_data.append({
                'wizard_id': self.id,
                'ps_name': product.name or '',
                'ps_internal_ref': product.default_code or '',
                'ps_category': product.categ_id.complete_name if product.categ_id else '',
                'ps_product_type': type_labels.get(product.type, product.type or ''),
                'ps_part_number_type': part_number_type,
                'ps_part_number': part_number,
                'ps_cost_price': product.standard_price,
                'ps_hsn_code': hsn_code,
                'ps_qty_on_hand': qty_on_hand,
            })
            row += 1

        workbook.close()
        output.seek(0)

        excel_data = base64.b64encode(output.read())
        filename = f'Product_Summary_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'

        self.preview_ids.unlink()
        self.env['report.preview'].create(preview_data)

        self.write({
            'excel_file': excel_data,
            'filename': filename,
        })

        view_id = self.env.ref('hrms_dashboard.view_reports_wizard_form').id
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'reports.wizard',
            'res_id': self.id,
            'view_mode': 'form',
            'views': [(view_id, 'form')],
            'target': 'current',
        }
