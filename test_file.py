from Invoicing import pdfInvoices
import os


pdfInvoices.generate('./test/test_invoices', './test/test_pdfs', './test/test_pythonhow.png', 'product_id', 'product_name', 'amount_purchased', 'price_per_unit', 'total_price')


