import yaml
import pyodbc
from xml.dom import minidom
import pandas


def read_yaml():
    with open('config.yaml') as config:
        return yaml.load(config, Loader=yaml.FullLoader)
    
def convert_xl_to_json(file_name):
    excel_data_df = pandas.read_excel(file_name)
    json_str = excel_data_df.to_json(orient='records')
    print('Excel Sheet to JSON:\n', json_str)

    
def parse_yaml(config_data):
    schema_table_dictionary = {}

    for mapper in config_data.get("rspbMapSmart"):
        for element, value in mapper.items():
            
             if "excelcolumn" in value:
                key = "excel"
                data_map_dictionary = {"column": value.get("excelcolumn"), "category": value.get("category"),
                                       "attribute": value.get("attribute"), "type": value.get("type")}
            else:
                key = value.get("schema") + "." + value.get("table")
                data_map_dictionary = {"column": value.get("column"), "category": value.get("category"),
                                   "attribute": value.get("attribute"), "type": value.get("type")}
            if key not in schema_table_dictionary:
                schema_table_dictionary[key] = []

            schema_table_dictionary[key].append(data_map_dictionary)

    return schema_table_dictionary


def create_select_queries(schema_table_dictionary):
    select_queries = []

    for schema_table, value in schema_table_dictionary.items():
        select_query = "select "

        index = 1
        for data_map_dictionary in value:
            select_query = select_query + data_map_dictionary["column"]

            if index != len(value):
                select_query = select_query + ", "

            index = index + 1

        select_query = select_query + " from " + schema_table

        select_queries.append(select_query)

    return select_queries


def run_queries(query):
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=DESKTOP-MSEQ9PA;'
                          'Database=master;'
                          'Trusted_Connection=yes;')

    cursor = conn.cursor()
    cursor.execute(query)

    #for i in cursor:
        #print(i)


def generate_smart_category_attribute_dic(schema_table_dictionary):
    attribute_by_category = {}
    for key, value in schema_table_dictionary.items():
        for element in value:
            cat = element['category']
            if cat not in attribute_by_category:
                attribute_by_category[cat] = []
            attribute = {}
            attribute['attribute'] = element['attribute']
            attribute['type'] = element['type']
            attribute_by_category[cat].append(attribute)

    return attribute_by_category


def generate_xml(attribute_by_category):
    root = minidom.Document()

    xml = root.createElement('DataModel')
    xml.setAttribute('xmlns', 'http://www.smartconservationsoftware.org/xml/1.0/datamodel')
    root.appendChild(xml)

    languages = root.createElement('languages')
    language = root.createElement('language')
    language.setAttribute('code', 'en')
    languages.appendChild(language)

    attributesElement = root.createElement('attributes')
    categories = root.createElement('categories')

    for category, attributes in attribute_by_category.items():
        categoryElement = root.createElement('category')
        categoryElement.setAttribute('ismultiple', 'true')
        categoryElement.setAttribute('isactive', 'true')
        categoryElement.setAttribute('key', category)

        categoryNamesElement = root.createElement('names')
        categoryNamesElement.setAttribute('language_code', 'en')
        categoryNamesElement.setAttribute('value', category)

        categoryElement.appendChild(categoryNamesElement)

        for attribute in attributes:
            attributeElement = root.createElement('attribute')
            attributeElement.setAttribute('key', attribute['attribute'])
            attributeElement.setAttribute('isrequired', 'false')
            attributeElement.setAttribute('type', attribute['type'])

            attributeNamesElement = root.createElement('names')
            attributeNamesElement.setAttribute('language_code', 'en')
            attributeNamesElement.setAttribute('value', attribute['attribute'])

            if attribute['type'] == 'TEXT':
                qa_regex = root.createElement('qa_regex')
                attributeElement.appendChild(qa_regex)

            attributeElement.appendChild(attributeNamesElement)
            attributesElement.appendChild(attributeElement)

            attributeCategoryElement = root.createElement('attribute')
            attributeCategoryElement.setAttribute('isactive', 'true')
            attributeCategoryElement.setAttribute('attributekey', attribute['attribute'])
            categoryElement.appendChild(attributeCategoryElement)

        categories.appendChild(categoryElement)

    xml.appendChild(languages)
    xml.appendChild(attributesElement)
    xml.appendChild(categories)

    return root.toprettyxml(indent="\t")


def save_xml(xml_str):
    save_path_file = "gfg.xml"

    with open(save_path_file, "w") as f:
        f.write(xml_str)


if __name__ == "__main__":
    config_data = read_yaml()
    schema_table_dictionary = parse_yaml(config_data)
    select_queries = create_select_queries(schema_table_dictionary)
    #for query in select_queries:
     #   run_queries(query)

    attribute_by_category = generate_smart_category_attribute_dic(schema_table_dictionary)
    xml_str = generate_xml(attribute_by_category)
    save_xml(xml_str)
