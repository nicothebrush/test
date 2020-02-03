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
def get_last_cost(raw_material_price, default_code, job_date, last_history):
    """ Extract last cost:
    """
    job_date = job_date[:10].replace('-', '')
    last = 0.0
    first = 0.0
    for date in sorted(raw_material_price[default_code]):
        if not first:
            first = raw_material_price[default_code][date]
            
        if job_date > date:
            last = raw_material_price[default_code][date]                        
        else:
            break
    comment = ''        
    if not last:
        last = last_history.get(default_code, 0.0)
        comment = 'Use history'

    if not last:
        comment = 'No cost'
        
    print comment, default_code, job_date, date, last
    return last
    
def get_cost(mrp, raw_material_price, current_cl, last_history):
    """ Get total for production closed
    """
    warning = []
    unload_document = []

    mrp_current_cost = current_cl.get(mrp.product_id.default_code, 0.0)
    
    cost_detail = u'' # To update MRP at the end of procedure
    cost_detail_subtotal = unload_cost_total = total = 0.0
    
    # Partial (calculated every load on all production)                
    cost_detail += u'Lavorazioni toccate:\n'
    cost_line_ref = u''
    wc = False
    for l in mrp.workcenter_lines:
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

            unload_document.append((
                partial.accounting_cl_code,
                partial.product_qty,
                l.real_date_planned[:10],
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
        warning.append('Not cost for line')

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
                last_cost = get_last_cost(
                    raw_material_price, 
                    unload.product_id.default_code,
                    lavoration.real_date_planned[:10],
                    last_history,
                    )
                subtotal = last_cost * unload.quantity
                unload_cost_total += subtotal
                cost_detail_subtotal += subtotal
                                    
                cost_detail += \
                    u' - %s: EUR %s x q. %s = %s\n' % (
                        unload.product_id.default_code or '?',
                        last_cost,
                        unload.quantity,
                        subtotal,
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
                        )
                    subtotal = last_cost * load.ul_qty
                    unload_cost_total += subtotal
                    cost_detail_subtotal += subtotal
                        
                cost_detail += \
                    u' - Imballo [%s] %s: EUR %s x q. %s = %s\n' % (
                        l.name,
                        link_product.default_code or '?',
                        last_cost,
                        load.ul_qty,
                        subtotal,
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
                        )
                    subtotal = last_cost * load.pallet_qty
                    unload_cost_total += subtotal
                    cost_detail_subtotal += subtotal
                    cost_detail += \
                        u' - Pallet [%s] %s: EUR %s x q. %s = %s \n' % (
                            l.name,
                            pallet_in.default_code or '?',
                            last_cost,
                            load.pallet_qty,
                            subtotal,
                            )
            except:
                warning.append('Error calculating pallet price')

    cost_detail += u'Totale imballi: %s\n' % cost_detail_subtotal
    if total:
        unload_cost = unload_cost_total / total
    else:
        unload_cost = 0.0    

    cost_detail += u'\nCosto totale:\n'
    cost_detail += u'EUR %s : q. %s = EUR/unit %s (carico)\n' % (
            unload_cost_total, total, unload_cost)

    for document in unload_document:
        print (
            document[0], 
            document[1], 
            document[2], 
            cost_detail, 
            mrp_current_cost, 
            unload_cost,
            )

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

last_history = {}
for line in open('ultimo.csv', 'r'):
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
for line in open('bfpan.csv', 'r'):
    line = line.strip()
    if not line:
        continue
    row = line.split(';')

    # Extract data:
    date = row[2].strip()
    default_code = row[5].strip()
    try:
        cost = float(row[8].strip().replace(',', '.'))
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
for line in open('cl2019.csv', 'r'):
    line = line.strip()
    if not line:
        continue
    row = line.split(';')

    # Extract data:
    default_code = row[5].strip()
    try:
        cost = float(row[8].strip().replace(',', '.'))
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

for mrp in mrp_pool.browse(mrp_ids):
    get_cost(mrp, raw_material_price, current_cl, last_history)
