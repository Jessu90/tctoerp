# -*- coding: utf-8 -*-
__author__ = 'Timo'

'''
Program which writes Teamcenter Item XML exports to Odoo products via XMLRPC.

Functionality:
Program polls  read_path, when .xml-file appears it will be read and sent to Odoo via XMLRPC.
Program writes logfile from executed product creations and exceptions.
Program writes timelog which is used to calculate offline time of program.
'''

import xmlrpclib, datetime, xml.etree.cElementTree
from time import clock, sleep
from sys import exit
from os import listdir, rename


# DEFINITIONS *********************************************************************************************************

# File locations.
read_path = 'C:/Temp/erp/in/'
archive_path = 'C:/Temp/erp/archieve/'
exception_path = 'C:/Temp/erp/exception/'
log_path = 'C:/Temp/erp/'
delay = 5  # Time interval [s] to poll read_path.

# Odoo server information.
usr = 'admin'
pwd = 'admin1'
db = 'RPC_1'
url = 'http://192.168.109.128:8069'

# RPC definitions.
common = xmlrpclib.ServerProxy(url+'/xmlrpc/2/common')
models = xmlrpclib.ServerProxy(url+'/xmlrpc/2/object')
uid = common.authenticate(db, usr, pwd, {})  # Asks Odoo user id of usr, used in queries besides username.


ET = xml.etree.cElementTree
DT = datetime.datetime
now_iso = DT.now().isoformat()  # Returns current time in str as isoformat.

def log(logstr):  # Create logfile.
    RPC_time = ' Time elapsed to RPC operations: %s s.' % (str(clock()-start_time)) if start_time else ''
    row = now_iso + '   ' + logstr + RPC_time
    print row
    file = open(log_path + 'log.txt', 'a')
    file.write(row + '\n')
    file.close()

def timelog():  # Def to update log used to calculate offline time.
    file = open(log_path+'timelog.txt', 'w')
    file.write(now_iso)
    file.close()

def call(obj, method, *parameters):  # RPC exception handling for Odoo server errors, unofrotunately connection errors are not handled.
    try:
        return models.execute_kw(db, uid, pwd, obj, method, *parameters)
    except Exception as e:
        log('ERROR: During task "%s", error: "%s".' % (oldest, e))
        exit(0)

def archive():  # Moves oldest (in use) XML file to archive.
    rename(read_path+oldest, archive_path+oldest)

def exception():  # Moves oldest (in use) XML file to exception-folder.
    rename(read_path+oldest, exception_path+oldest)

def find_product_product_ids():  # Def to find dbIDs of product_product in Odoo which are consumed through BOM lines. Returns IDs as list.
    restart_cycle = False
    product_product_ids = []
    for mrp_bom_line in mrp_bom_lines:
        product_product_id = call('product.product', 'search', [[['default_code', '=', mrp_bom_line['product_id']]]])  # Returns ID's of products consumed on BOM lines.
        if not product_product_id:  # If consumed product does not exist in Odoo, exception.
            log('EXCEPTION: Task: "%s". Products to be consumed through BOM lines of product "%s" does not exist in Odoo. Atleast product "%s" is missing.' % (oldest, product_template['default_code'], mrp_bom_line['product_id']))
            exception()
            restart_cycle = True
            break
        elif len(product_product_id) != 1:
            log('EXCEPTION: Task: "%s". There is already multiple instances (%s) of product "%s" to consume through BOM line in Odoo.' % (oldest, len(product_product_id), mrp_bom_line['product_id']))
            exception()
            restart_cycle = True
            break
        product_product_ids.append(product_product_id[0])
    if restart_cycle == True:
        return None
    else:
        return product_product_ids

def create_mrp_bom_lines():  # Def to create new BOM lines to model mrp_bom_line.
    index_nro = 0
    created_mrp_bom_line_ids = []  # dbID list of BOM lines which will be created.
    for mrp_bom_line in mrp_bom_lines:  # Creation of BOM lines.
        mrp_bom_line['bom_id'] = mrp_bom_id
        mrp_bom_line['product_id'] = product_ids[index_nro]
        index_nro += 1
        mrp_bom_line_id = call('mrp.bom.line', 'create', [mrp_bom_line])  # Creation of BOM line.
        created_mrp_bom_line_ids.append(mrp_bom_line_id)
    return created_mrp_bom_line_ids



# OFFLINE TIME CHECK **************************************************************************************************

try:  # Open timelog to read last timestamp.
    timelogfile = open(log_path+'timelog.txt', 'r')
    from_timelog = timelogfile.read()
except IOError:  # If timelog does not exist it will be created.
    timelogfile = open(log_path+'timelog.txt', 'w')
    from_timelog = None
finally:  # Anyway timelog will be closed.
    timelogfile.close()

start_time = None  # Definition of needed variable, defined again later.
if from_timelog:  # If there was old timelog it will be used in calculation of offline time elapsed from last log.
    logtime = DT.strptime(from_timelog, "%Y-%m-%dT%H:%M:%S.%f" )  # str time to DT format.
    elapsed = DT.now()-logtime
    log('START: Last online time was: %s, offline time after that: %s.' % (from_timelog, str(elapsed)))
else:
    log('START: Old timestamp does not found.')

timelog()  # Update timelog.


while True:

    # PATH LISTENING ******************************************************************************************************

    filenames = listdir(read_path)  # Lists file names of files in path.

    row_nro = 0
    while not filenames:  # If there is not files in path, wait.
        print'%s. No files in path: "%s", idle time %s seconds.' % (row_nro, read_path, delay)
        timelog()
        sleep(delay)
        row_nro += 1
        filenames = listdir(read_path)

    oldest = filenames[0] if filenames else None  # Return the name of oldest file in the path (regarding filename).



    # XML FILE OPERATIONS **************************************************************************************************

    # Form XML tree from oldest file.
    tree = ET.parse(read_path+oldest)
    root = tree.getroot()

    # Take required elements from XML file.
    Items = root.find('Items')
    Item = Items.find('./Item[@isTopLevel="Yes"]')
    Attributes = Item.find('AdditionalAttributes')
    BillOfMaterial = Item.find('BillOfMaterial')

    # Take values from XML file.
    ID = Item.attrib['itemIdentifier']
    Revision = Item.attrib['revisionIdentifier']
    Name = Attributes.find('./AdditionalAttribute[@name="object_name"]').attrib['value']
    Description = Attributes.find('./AdditionalAttribute[@name="sea3Description"]').attrib['value']  # From Item.
    Product = Attributes.find('./AdditionalAttribute[@name="sea3Product"]').attrib['value']
    RelatedTo = Attributes.find('./AdditionalAttribute[@name="sea3RelatedTo"]').attrib['value']
    Mass = Attributes.find('./AdditionalAttribute[@name="sea3Mass"]').attrib['value']
    Designer = Attributes.find('./AdditionalAttribute[@name="sea3Designer"]').attrib['value']
    ManufacturerOrSupplier = Attributes.find('./AdditionalAttribute[@name="sea3ManufacturerOrSupplier"]').attrib['value']
    MatOrStandardCode = Attributes.find('./AdditionalAttribute[@name="sea3MatOrStdOrCode"]').attrib['value']
    MaterialSize = Attributes.find('./AdditionalAttribute[@name="sea3SizeOfMaterial"]').attrib['value']
    Info = Attributes.find('./AdditionalAttribute[@name="sea3info"]').attrib['value']
    SparePart = Attributes.find('./AdditionalAttribute[@name="sea3Spare_part"]').attrib['value']
    ChangeDescription = Attributes.find('./AdditionalAttribute[@name="sea3ChangeDescription"]').attrib['value']



    # ODOO DATA TEMPLATE FILLING FROM XML***********************************************************************************

    # Fill product_template.
    product_template = {
        'name' : Name,
        'description' : 'Active revision (in use): %s. \n'\
                        'Description: %s. \n'
                        'Product: %s. \n'
                        'Related to: %s. \n'
                        'Designer: %s \n'
                        'Change description (between revisions): %s. \n'
                        'Sold as spare part: %s.'
                        % (Revision, Description, Product, RelatedTo, Designer, ChangeDescription, SparePart),
        'description_purchase' : 'Manufacturer or supplier: %s. \n'
                        'Material code or standard: %s. \n'
                        'Material size: %s. \n'
                        'Info: %s. \n'
                        % (ManufacturerOrSupplier, MatOrStandardCode, MaterialSize, Info),
        'weight_net' : '0.0' if Mass == '' else round(float(Mass), 3),
        'type' : 'product',  # Stockable product.
        'state' : 'draft',  # State: In Development.
        'default_code' : ID,  # TCID of product, stored to product.product.
        'warranty'  : 12.00
    }


    if BillOfMaterial:  # Fill  BOM and list of BOM lines if item does have a BOM. BOM data will be stored to mrp.bom and mrp.bom.line.
        mrp_bom = {
        'name' : Name,
        'code' : ID,  # Product's TCID will be stored to BOM reference too.
        'product_tmpl_id' : '', # dbID of product_template which uses BOM, filled afterward.
        'sequence' : '0'  # Order of BOM, default is 0.
        }

        mrp_bom_lines = []
        for BillOfMaterialItem in BillOfMaterial:  # Do a list from BOM lines in XML.
            BOMline = BillOfMaterialItem.attrib
            mrp_bom_line = {
            'bom_id' : '',  # ID of BOM which contains this BOM line, filled afterward.
            'sequence' : BOMline['BOM_SequenceNumber'],  # Order nro of BOM line.
            'product_id' : BOMline['BOM_ItemID'],  # TCID of the product which is consumed through this BOM, changed to dbID afterward.
            'product_qty' : BOMline['BOM_itemQuantity'],  # Amount of products consumed on this line.
            'type' : 'normal'  # normal = no Phantom item.
            }
            mrp_bom_lines.append(mrp_bom_line)



    # RPC OPERATIONS *******************************************************************************************************

    start_time = clock() # Timer to evaluate RPC performance.

    # Check if Product is already in Odoo. Take into account that variant of product doesn't share it's parent default_code.
    product_product_exist = call('product.product', 'search_read', [[['default_code', '=', product_template['default_code']]]], {'fields': ['product_tmpl_id']})
    product_product_id = product_product_exist[0]['id'] if product_product_exist else None
    product_tmpl_id = product_product_exist[0]['product_tmpl_id'][0] if product_product_exist else None



    if not product_product_exist:  # If product does not exist in Odoo, new product will be created.

        if BillOfMaterial:  # If product does have a BOM, check that products to consume exist in Odoo.
            product_ids = find_product_product_ids()  # dbID list of products to consume through BOM lines.
            if not product_ids:  # If exception happened when searching products to consume return to begin of While True:.
                continue

        product_tmpl_id = call('product.template', 'create', [product_template])  # New product_template and product_product will be created.

        if BillOfMaterial:  # If product does have a BOM. It will be created.
            mrp_bom['product_tmpl_id'] = product_tmpl_id
            mrp_bom_id = call('mrp.bom', 'create', [mrp_bom])  # BOM creation.

            created_bom_line_ids = create_mrp_bom_lines()  # BOM lines creation.

            log('CREATE: Product "%s" created at product_tmpl_id: "%s" and BOM created at mrp_bom_id: "%s" with mrp_bom_line_ids: "%s".' % (product_template['default_code'], product_tmpl_id, mrp_bom_id, created_bom_line_ids))
            archive()

        else:  # If product does not have a BOM, this log entry will be used with product creation.
            log('CREATE: Product "%s" created at product_tmpl_id: "%s".' % (product_template['default_code'], product_tmpl_id))
            archive()




    elif len(product_product_exist) == 1: # If product already exist in Odoo it will be updated.

        if BillOfMaterial:  # If product does have a BOM, check conditions to update it.
            mrp_bom_id = call('mrp.bom', 'search', [[['product_tmpl_id', '=', product_tmpl_id]]])  # Find BOMs which are referring to product in Odoo.
            if len(mrp_bom_id) != 1:  # If multiple BOMs referring to product, exception.
                log('EXCEPTION: Task: "%s". Unexcepted amount (%s) of BOMs referring to product "%s" in Odoo. Problematic mrp_bom_ids: "%s".' % (oldest, len(mrp_bom_id), product_template['default_code'], mrp_bom_id))
                exception()
                continue
            mrp_bom_id = mrp_bom_id[0]

            product_ids = find_product_product_ids()  # dbID list of products to consume through BOM lines.
            if not product_ids:  # If exception happened when searching products to consume return to begin of While True:.
                continue

        call('product.template', 'write', [product_tmpl_id, product_template])  # Update product_template. Returns True.

        if BillOfMaterial:  # If product does have a BOM, it will be updated.
            # Update existing BOM.
            mrp_bom['product_tmpl_id'] = product_tmpl_id
            call('mrp.bom', 'write', [mrp_bom_id, mrp_bom])  # BOM update.

            # Deletion of existing BOM lines.
            mrp_bom_line_ids = call('mrp.bom.line', 'search', [[['bom_id', '=', mrp_bom_id]]])  # Find BOM lines which are referring to BOM.
            if mrp_bom_line_ids:  # Only executes BOM line deletion when there is existing BOM lines.
                call('mrp.bom.line', 'unlink', [mrp_bom_line_ids])  # Deletion of existing BOM lines (list of dbIDs) which are referring to BOM which is to be updated.

            # Creation of new BOM lines.
            created_bom_line_ids = create_mrp_bom_lines()

            log('UPDATE: Product "%s" updated, product.product id: "%s", product.template id: "%s". BOM updated at mrp_bom_id: "%s" with mrp_bom_line_ids: "%s".' % (product_template['default_code'], product_product_id, product_tmpl_id, mrp_bom_id, created_bom_line_ids))
            archive()

        else:  # If product does not have a BOM, this log entry will be used with product update.
            log('UPDATE: Product "%s" updated, product.product id: "%s", product.template id: "%s".' % (product_template['default_code'], product_product_id, product_tmpl_id))
            archive()




    else:  # If multiple instances of TC product (ID) already in Odoo --> Exception.
        log('EXCEPTION: Task: "%s". There is already multiple instances (%s) of product "%s" in Odoo. Problematic products: "%s".' % (oldest, len(product_product_exist), product_template['default_code'], product_product_exist))
        exception()
