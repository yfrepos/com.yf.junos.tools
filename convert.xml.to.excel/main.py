import xml.etree.ElementTree as ET
import pandas as pd
import os

def process_xml_file(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    addresses = []
    address_sets = []
    applications = []
    application_set = []
    policies = []

    excel_file = os.path.splitext(xml_file)[0] + '.xlsx'
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

    for root_item in root:
        if root_item.tag == 'configuration':
            for conf_item in root_item:
                if conf_item.tag == 'security':
                    for sec_item in conf_item:
                        if sec_item.tag == 'address-book':
                            for address_book_item in sec_item:
                                if address_book_item.tag == 'address':
                                    name = ''
                                    description = ''
                                    ip_prefix = ''

                                    for address in address_book_item:
                                        if address.tag == 'name':
                                            name = address.text
                                        if address.tag == 'description':
                                            description = address.text
                                        if address.tag == 'ip-prefix':
                                            ip_prefix = address.text
                                        if address.tag == 'range-address':
                                            range_start = address.find('name').text
                                            range_end = address.find('to/range-high').text
                                            ip_prefix = f"{range_start}-{range_end}"

                                    addresses.append([name, description, ip_prefix])

                                if address_book_item.tag == 'address-set':
                                    name = ''
                                    description = ''
                                    addresses_in_set = []

                                    for address_set in address_book_item:
                                        if address_set.tag == 'name':
                                            name = address_set.text
                                        if address_set.tag == 'description':
                                            description = address_set.text
                                        if address_set.tag == 'address':
                                            addresses_in_set.append(address_set[0].text)

                                    address_sets.append([name, description, ';'.join(addresses_in_set)])

                        if sec_item.tag == 'policies':
                            for policy in sec_item:
                                if policy.tag == 'policy':
                                    for policy_item in policy:
                                        status = 'active'
                                        if 'inactive' in policy_item.attrib:
                                            status = 'inactive'

                                        if policy_item.tag == 'from-zone-name':
                                            from_zone_name = policy_item.text
                                        if policy_item.tag == 'to-zone-name':
                                            to_zone_name = policy_item.text

                                        subpol_description = ''
                                        sub_match_source_addr = []
                                        sub_match_dest_addr = []
                                        sub_match_app = []
                                        action = ''

                                        if policy_item.tag == 'policy':
                                            for subpolicy in policy_item:
                                                if subpolicy.tag == 'name':
                                                    subpol_name = subpolicy.text
                                                if subpolicy.tag == 'description':
                                                    subpol_description = subpolicy.text or ''
                                                if subpolicy.tag == 'match':
                                                    for match in subpolicy:
                                                        if match.tag == 'source-address':
                                                            sub_match_source_addr.append(match.text)
                                                        if match.tag == 'destination-address':
                                                            sub_match_dest_addr.append(match.text)
                                                        if match.tag == 'application':
                                                            sub_match_app.append(match.text)
                                                if subpolicy.tag == 'then':
                                                    for then in subpolicy:
                                                        if then.tag == 'permit':
                                                            action = 'permit'
                                                        if then.tag == 'deny':
                                                            action = 'deny'

                                            policies.append([subpol_name, status, from_zone_name, to_zone_name, 
                                                             ';'.join(sub_match_source_addr), 
                                                             ';'.join(sub_match_dest_addr), 
                                                             ';'.join(sub_match_app), 
                                                             subpol_description, action])

                if conf_item.tag == 'applications':
                    for application in conf_item:
                        if application.tag == 'application':
                            name = ''
                            source_port = ''
                            dest_port = ''
                            protocol = ''
                            term_label = []
                            for application_item in application:
                                if application_item.tag == 'name':
                                    name = application_item.text
                                if application_item.tag == 'protocol':
                                    protocol = application_item.text
                                if application_item.tag == 'destination-port':
                                    dest_port = application_item.text
                                if application_item.tag == 'source-port':
                                    source_port = application_item.text
                                if application_item.tag == 'term':
                                    term_label_dest_port = ''
                                    term_label_protocol = ''
                                    for term in application_item:
                                        if term.tag == 'destination-port':
                                            term_label_dest_port = term.text
                                        if term.tag == 'protocol':
                                            term_label_protocol = term.text

                                        if term_label_dest_port and term_label_protocol:
                                            term_label.append(f"{term_label_dest_port}/{term_label_protocol}")

                            applications.append([name, source_port, dest_port, protocol, 
                                                 ';'.join(term_label)])

                        if application.tag == 'application-set':
                            name = ''
                            application_in_set = []
                            for app_set_item in application:
                                if app_set_item.tag == 'name':
                                    name = app_set_item.text
                                if app_set_item.tag == 'application':
                                    application_in_set.append(app_set_item[0].text)

                            application_set.append([name, ';'.join(application_in_set)])

    if addresses:
        df1 = pd.DataFrame(addresses, columns=['Name', 'Description', 'IP Address'])
        df1.to_excel(writer, index=False, sheet_name='Addresses')

    if address_sets:
        df2 = pd.DataFrame(address_sets, columns=['Name', 'Description', 'Addresses'])
        df2.to_excel(writer, index=False, sheet_name='Address-sets')

    if policies:
        df3 = pd.DataFrame(policies, columns=['name', 'status', 'from-zone-name', 'to-zone-name', 'source-address', 'destination-address', 'application', 'description', 'action'])
        df3.to_excel(writer, index=False, sheet_name='Policies')

    if applications:
        df4 = pd.DataFrame(applications, columns=['Name', 'Source port', 'Destination port', 'Protocol', 'Destination ports/protocol'])
        df4.to_excel(writer, index=False, sheet_name='Applications')

    if application_set:
        df5 = pd.DataFrame(application_set, columns=['Name', 'Applications'])
        df5.to_excel(writer, index=False, sheet_name='Application-sets')

    writer.close()

def main():
    xml_files = [f for f in os.listdir() if f.endswith('.xml')]
    for xml_file in xml_files:
        process_xml_file(xml_file)
    print("Finished")

if __name__ == '__main__':
    main()
