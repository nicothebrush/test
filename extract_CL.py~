# -*- coding: utf-8 -*-
###############################################################################
#
# ODOO (ex OpenERP) 
# Open Source Management Solution
# Copyright (C) 2001-2015 Micronaet S.r.l. (<http://www.micronaet.it>)
# Developer: Nicola Riolini @thebrush (<https://it.linkedin.com/in/thebrush>)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. 
# See the GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.
#
###############################################################################
import os
import erppeek
import ConfigParser
import xlsxwriter

# ---------------------------------------------------------------------
# Export XLSX file:
# ---------------------------------------------------------------------
demo = False

# Excel writing:     
def xls_write_row(ws_name, row, row_data, row_format=None):
    ''' Write line in excel file
    '''
    col = 0
    for item in row_data:
        WS[ws_name].write(row, col, item, row_format)
        col += 1
    return True

def xls_row_width(ws_name, width_list):
    ''' Write line in excel file
    '''
    for col in range(len(width_list)):
        WS[ws_name].set_column(col, col, width_list[col])
    return True

def extract_price_mrp(mrp):
    """ Extract raw material from log 
    """
    res = {}
    cost_detail = mrp.cost_detail
    if cost_detail:
        cost_detail = cost_detail.replace('<br/>', '\n')
        for line in cost_detail.split('\n'):
            if line.startswith(' - '):
                default_code = line[3:].split(':')[0]
                cost = float((line[3:].split(':')[-1].split(
                    'x')[0].strip().split(' ')[-1]))
                
                res[default_code] = cost
    return res

# Open file and write header
file_out = 'log.xlsx'
WB = xlsxwriter.Workbook(file_out)

# Format setup:
xls_format = {
    'header': WB.add_format({
        'bold': True,
        'font_name': 'Verdana',
        'font_size': 9,
        'align': 'center',
        'bg_color': 'EEEEEE',
        'border': 1,
        #'num_format': '0.00',
        }),
    'text': WB.add_format({
        'font_name': 'Verdana',
        'font_size': 9,
        'align': 'center',
        #'bg_color': 'EEEEEE',
        'border': 1,
        #'num_format': '0.00',
        }),
    'number': WB.add_format({
        'font_name': 'Verdana',
        'font_size': 9,
        'align': 'center',
        #'bg_color': 'EEEEEE',
        'border': 1,
        'num_format': '0.00',
        }),
        
    }

# Worksheet:
WS = {
    'Costo': WB.add_worksheet('Costo'),
    'Ultimo': WB.add_worksheet('Ultimo'),
    'Senza': WB.add_worksheet('Senza'),
    }
counter = {
    'Costo': 1,
    'Ultimo': 1, 
    }    

# Header Costo
xls_write_row('Costo', 0, (
    'CL', 'Q.', 'Prodotto',
    'MRP', '#',
    'MRP scar.', 'MRP car.', 'Diff.', 'Stato',    
    'Data', 'Detail', 'ODOO Detail',
    'Mexal', 'ODOO', 'Diff.', 'Status',
    'Warning',
    ), xls_format['header'])

xls_row_width('Costo', [
    10, 10, 18,
    10, 2, 
    10, 10, 8, 5, 
    10, 40, 40,
    15, 15, 15, 5,
    50,
    ])

# Header Ultimo    
xls_write_row('Ultimo', 0, (
    'Commento',
    'Codice',
    'Data lav.',
    'Data BF',
    'Costo',
    ), xls_format['header'])

xls_row_width('Ultimo', [
    30, 15, 10, 10, 12,
    ])

empty_cost = []
# -----------------------------------------------------------------------------
# Read configuration parameter:
# -----------------------------------------------------------------------------
cfg_file = os.path.expanduser('../openerp.cfg')
config = ConfigParser.ConfigParser()
config.read([cfg_file])
dbname = config.get('dbaccess', 'dbname')
user = config.get('dbaccess', 'user')
pwd = config.get('dbaccess', 'pwd')
server = config.get('dbaccess', 'server')
port = config.get('dbaccess', 'port')   # verify if it's necessary: getint

# -----------------------------------------------------------------------------
#                                   UTILITY:
# -----------------------------------------------------------------------------
def get_last_cost(
        raw_material_price, default_code, job_date, last_history, mrp_cost,
        odoo_standard):
    """ Extract last cost:
    """
    job_date = job_date[:10].replace('-', '')
    last = first = 0.0
    date = 'Non trovata'

    comment = ''
    if default_code.startswith('VV'):
        comment = 'Not considered'        
    else:    
        for date in sorted(raw_material_price.get(default_code, [])):
            if not first:
                first = raw_material_price[default_code][date]
                
            if job_date > date:
                last = raw_material_price[default_code][date]                        
            else:
                break

        if not last:
            last = last_history.get(default_code, 0.0)
            comment = 'Use Mexal last cost'

        if not last:
            last = mrp_cost.get(default_code, 0.0)
            comment = 'Use ODOO MRP detail'
            
        if not last:
            last = odoo_standard.get(default_code, 0.0)
            comment = 'Use ODOO standard cost'
            
        if not last:
            if default_code not in empty_cost:
                empty_cost.append(default_code)
            comment = 'No cost present'
        
    row = counter['Ultimo']
    counter['Ultimo'] += 1
    xls_write_row(
        'Ultimo', row, 
        (comment, default_code, job_date, date, last),
        xls_format['text'])
    return last
    
def get_cost(mrp, raw_material_price, current_cl, last_history, odoo_standard):
    """ Get total for production closed
    """
    warning = []
    unload_document = []
    
    mrp_cost = extract_price_mrp(mrp)  # Extract cost from mrp detail
    mrp_code = mrp.product_id.default_code

    mrp_current_cost = current_cl.get(mrp_code, 0.0)
    
    cost_detail = u'' # To update MRP at the end of procedure
    cost_detail_subtotal = unload_cost_total = total = total_unload = 0.0
    
    # Partial (calculated every load on all production)                
    cost_detail += u'Lavorazioni toccate:\n'
    cost_line_ref = u''
    wc = False
    for l in mrp.workcenter_lines:
        if l.state != 'done':
            print '%s Not in done state' % l.name
            continue

        if not wc:
            wc = l.workcenter_id
        for partial in l.load_ids:

            if not cost_line_ref:
                cost_line_ref = l.workcenter_id.name or '?'
            total += partial.product_qty or 0.0
            cost_detail += u' [%s q.: %s]' % (
                l.name,
                partial.product_qty,
                )

    # Loop for print part (for total purpose)
    for l in mrp.workcenter_lines:
        for partial in l.load_ids:
            unload_document.append((
                partial.accounting_cl_code,
                partial.product_qty,
                l.real_date_planned[:10],
                partial.accounting_cl_code,
                total,
                ))

    cost_detail += u'\nTotale carichi: %s\n' % total

    # -------------------------------------------------------------------------
    # Lavoration K cost:
    # -------------------------------------------------------------------------
    try:
        cost_line = wc.cost_product_id.standard_price or 0.0
    except:
        print 'No line cost'
        cost_line = 0.0
    
    if not cost_line:    
        warning.append('Line K not found!')

    unload_cost_total = cost_line * total
    cost_detail += u'\nLavorazione %s: euro/kg. %s x %s = %s\n' % (
            cost_line_ref, cost_line, total, unload_cost_total)

    # -------------------------------------------------------------------------
    # All unload cost of materials (all production):
    # -------------------------------------------------------------------------
    cost_detail += \
        u'\nCosti materie prime:\n'
    cost_detail_subtotal = 0.0
    for lavoration in mrp.workcenter_lines:
        cost_detail += u'Lavoration %s:\n' % lavoration.name
        for unload in lavoration.bom_material_ids:
            try:
                default_code = unload.product_id.default_code or '?'

                last_cost = get_last_cost(
                    raw_material_price, 
                    unload.product_id.default_code,
                    lavoration.real_date_planned[:10],
                    last_history,
                    mrp_cost,
                    odoo_standard,
                    )
                if default_code[:2] != 'VV' and not last_cost:
                    warning.append('Material with price 0')
    
                total_unload += unload.quantity
                subtotal = last_cost * unload.quantity
                unload_cost_total += subtotal
                cost_detail_subtotal += subtotal
                                    
                cost_detail += \
                    u' - %s: EUR %s x q. %s = %s %s\n' % (
                        default_code,
                        last_cost,
                        unload.quantity,
                        subtotal,
                        '***' if not last_cost else '',
                        )
            except:
                warning.append('Error calculating unload lavoration')                
    cost_detail += u'Totale materie prime: %s\n' % cost_detail_subtotal

    # -------------------------------------------------------------------------
    # All unload package and pallet:
    # -------------------------------------------------------------------------
    cost_detail += u'\nCosti imballi e pallet:\n'
    for l in mrp.workcenter_lines:                    
        for load in l.load_ids:
            try:
                # -------------------------------------------------------------
                # Package:
                # -------------------------------------------------------------
                package = load.package_id                                
                if package: # There's pallet
                    link_product = package.linked_product_id
                    last_cost = get_last_cost(
                        raw_material_price, 
                        link_product.default_code,
                        load.date,
                        last_history,
                        mrp_cost,
                        odoo_standard,
                        )
                    if not last_cost:
                        warning.append('Package with price 0')
                        
                    subtotal = last_cost * load.ul_qty
                    unload_cost_total += subtotal
                    cost_detail_subtotal += subtotal
                        
                cost_detail += \
                    u' - Imballo [%s] %s: EUR %s x q. %s = %s %s\n' % (
                        l.name,
                        link_product.default_code or '?',
                        last_cost,
                        load.ul_qty,
                        subtotal,
                        '***' if not last_cost else '',
                        )
            except:
                warning.append('Error calculating package price')
                
            try:
                # -------------------------------------------------------------
                # Pallet:
                # -------------------------------------------------------------
                pallet_in = load.pallet_product_id
                if pallet_in: # there's pallet
                    last_cost = get_last_cost(
                        raw_material_price, 
                        pallet_in.default_code,
                        load.date,
                        last_history,
                        mrp_cost,
                        odoo_standard,
                        )
                    if not last_cost:
                        warning.append('Pallet with price 0')

                    subtotal = last_cost * load.pallet_qty
                    unload_cost_total += subtotal
                    cost_detail_subtotal += subtotal
                    cost_detail += \
                        u' - Pallet [%s] %s: EUR %s x q. %s = %s %s\n' % (
                            l.name,
                            pallet_in.default_code or '?',
                            last_cost,
                            load.pallet_qty,
                            subtotal,
                            '***' if not last_cost else '',
                            )
            except:
                warning.append('Error calculating pallet price')

    cost_detail += u'Totale imballi: %s\n' % cost_detail_subtotal
    if total:
        unload_cost = unload_cost_total / total
    else:
        unload_cost = 0.0    

    cost_detail += u'\nPeso materie prime: %s\n' % total_unload
    cost_detail += u'\nCosto totale:\n'
    cost_detail += u'EUR %s : q. %s = EUR/unit %s (carico)\n' % (
            unload_cost_total, total, unload_cost)

    res = set()    
    for document in unload_document:        
        if not mrp_current_cost:
            print 'No Mexal'
            continue    

        res.add(document[3]) # CL code

        # Difference status:
        difference = mrp_current_cost - unload_cost
        if abs(difference) <= 0.03:
            status = ''
        elif abs(difference) <= 1.0:
            status = 'X'
        elif abs(difference) <= 10.0:
            status = 'XX'
        elif abs(difference) <= 100.0:
            status = 'XXX'
        elif abs(difference) <= 1000.0:
            status = 'XXXX'
        else:    
            status = 'XXXXX'            
        
        # Weight status:
        weight_difference = total_unload - document[4]
        if (abs(weight_difference) / document[4]) > 0.1:
            weight_status = 'X'
        else:
            weight_status = ''

        # Counter:
        row = counter['Costo']
        counter['Costo'] += 1

        # Clean original detail:
        odoo_cost_detail = (mrp.cost_detail or '').replace(
            '<br/>', '\n').replace('<b>', '\n').replace('</b>', '\n')
            
        # Write line:
        xls_write_row('Costo', row, (        
            document[0], # CL
            document[1], # Q. 
            mrp_code,

            mrp.name,
            len(unload_document), # Number of CL

            total_unload, # Q. unload
            document[4], # MRP total
            weight_difference, 
            weight_status,
            
            document[2], # Date
            cost_detail, # Detail
            odoo_cost_detail, # ODOO detail
            mrp_current_cost, # Mexal
            unload_cost, # ODOO
            mrp_current_cost - unload_cost,
            status,
            ', '.join(warning),
            ), xls_format['text'])
        print row, document[0], document[1], unload_cost
    return res    

# -----------------------------------------------------------------------------
# Connect to ODOO:
# -----------------------------------------------------------------------------
odoo = erppeek.Client(
    'http://%s:%s' % (
        server, port), 
    db=dbname,
    user=user,
    password=pwd,
    )
print 'Connect with ODOO: %s' % odoo    

# -----------------------------------------------------------------------------
# Standard cost from ODOO
# -----------------------------------------------------------------------------
odoo_standard = {}
for line in open('./data/odoo_standard.csv', 'r'):
    line = line.strip()
    if not line:
        continue
    row = line.split('|')

    # Extract data:
    default_code = row[0].strip()
    try:
        cost = float(row[1].strip().replace(',', '.'))
    except:
        cost = 0.0
        print 'No last cost: %s' % line
    odoo_standard[default_code] = cost

# -----------------------------------------------------------------------------
# Last cost from Mexal
# -----------------------------------------------------------------------------
last_history = {}
for line in open('./data/cuppan.csv', 'r'):
    line = line.strip()
    if not line:
        continue
    row = line.split(';')

    # Extract data:
    default_code = row[0].strip()
    try:
        cost = float(row[1].strip().replace(',', '.'))
    except:
        cost = 0.0
        print 'No last cost: %s' % line
    last_history[default_code] = cost

# -----------------------------------------------------------------------------
# Load Last cost:
# -----------------------------------------------------------------------------
raw_material_price = {}
for filename in ('./data/bfpan18.csv', './data/bfpan19.csv'):
    for line in open(filename, 'r'):
        line = line.strip()
        if not line:
            continue
        row = line.split(';')

        # Extract data:
        date = row[2].strip()
        default_code = row[5].strip()
        try:
            cost = float(row[9].strip().replace(',', '.'))
        except:
            cost = 0.0
            print 'No price: %s' % line
        
        if default_code not in raw_material_price:
            raw_material_price[default_code] = {}
        raw_material_price[default_code][date] = cost

# -----------------------------------------------------------------------------
# Load current CL status
# -----------------------------------------------------------------------------
current_cl = {}
cl_mexal = set()
cl_odoo = set()

for line in open('./data/clpan19.csv', 'r'):
    line = line.strip()
    if not line:
        continue
    row = line.split(';')

    # Extract data:
    number = cl_mexal.add(row[1].strip())
    default_code = row[5].strip()
    try:
        cost = float(row[9].strip().replace(',', '.'))
    except:
        cost = 0.0
        print 'CL No price: %s' % line
    current_cl[default_code] = cost
    
# -----------------------------------------------------------------------------
# Check production
# -----------------------------------------------------------------------------
mrp_pool = odoo.model('mrp.production')

i = 0
mrp_ids = mrp_pool.search([
    ('date_planned', '>=', '2019-01-01'),
    ])
if demo:
    mrp_ids = mrp_ids[:2]
    
for mrp in mrp_pool.browse(mrp_ids):
    cl = get_cost(mrp, raw_material_price, current_cl, last_history, 
    odoo_standard)
    cl_odoo.union(cl)

print 'Differenza ODOO - Mexal', cl_odoo.difference(cl_mexal)
print 'Differenza Mexal - ODOO', cl_mexal.difference(cl_odoo)

row = -1

for empty in empty_cost:
    row += 1 
    xls_write_row('Senza', row, (
        empty,
        ), xls_format['text'])
    
WB.close()    
