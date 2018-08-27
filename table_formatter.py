from jinja2 import Markup
from dateutil import tz
import datetime
import pyodbc, csv

class TableFormatter(object):
    header_mapping = {}
    value_lookups = {}
    title = None
    source_tz=None
    display_tz=None
    timestamp_format='%Y-%m-%d %H:%M:%S'

    def __init__(self,table_data,field_names,header_mapping=None,value_lookups=None,title=None, source_tz=None, display_tz=None, timestamp_format=None):
        # Override class defaults with instance attributes
        self.header_mapping = header_mapping or self.header_mapping
        self.value_lookups = value_lookups or self.value_lookups
        self.title = title or self.title
        self.source_tz=source_tz or self.source_tz
        self.display_tz = display_tz or self.display_tz
        self.timestamp_format = timestamp_format or self.timestamp_format
        # Build translated headers
        self.headers=[(self.header_mapping[field_name]
            if field_name in self.header_mapping
            else field_name)
            for field_name in field_names]
        #Valdiate and process table data
        self.data = self.build_table_data(table_data,field_names)

    def build_table_data(self,table_data,field_names):
        # Make sure table data is valid format and convert into list of dicts
        if type(table_data) in {tuple,list}:
            if table_data:
                if type(table_data[0]) in {list, tuple}: #add  pyodbc.Row if required:
                    labelled_table = [dict(zip(field_names,row)) for row in table_data]
                elif isinstance(table_data[0],dict):
                    labelled_table=table_data
                else:
                    raise Exception('Table rows are invalid format (must be dict, list or tuple)')
            else:
                labelled_table=[]
        else:
            raise Exception('Table structure must be a list or tuple')

        # Convert table to list of lists with fields in order of field names and perform value lookups where necessary
        processed_table= [[self.translate_field_value(field,row[field]) for field in field_names]
                for row in labelled_table]
        return processed_table

    def translate_field_value(self,field,value):
        if isinstance(value,datetime.datetime):
            if self.source_tz and self.display_tz:
                value=value.replace(tzinfo=self.source_tz).astimezone(self.display_tz)
            return value.strftime(self.timestamp_format)
        elif field in self.value_lookups and value in self.value_lookups[field]:
            return self.value_lookups[field][value]
        else:
            return value

    # Format table as HTML with a maximum number of rows and optional classes
    def as_html(self, maxrows=20, classes='table-striped'):
        table_str = "<h6>{}</h6>".format(self.title) if self.title else ''
        if not self.data:
            table_str+='<p>No results</p>'
            return Markup(table_str)
        endstr = ''
        table_str += "<table class='table'><thead class='thead-light {}'><tr>".format(classes)
        # Add headers
        for header in self.headers:
            table_str += "<th>{}</th>".format(header)
        table_str +="</tr></thead><tbody>"
        # Add data
        for num, row in enumerate(self.data):
            table_str += "<tr>"
            for val in row:
                table_str += "<td>{}</td>".format(val if val is not None else '')
            table_str+="</tr>\n"
            if num >= maxrows:
                endstr='<p>Total {} results, only showing {}</p>'.format(len(self.data),maxrows)
                break
        # Finish
        table_str+="</tbody></table>" + endstr
        return Markup(table_str)

    # Write table to CSV
    def to_csv(self,file_path):
        with open(file_path,'wt') as csvfile:
            csvwriter=csv.writer(csvfile)
            csvwriter.writerow(self.headers)
            csvwriter.writerows(self.data)

    def __call__(self,**kwargs):
        return self.as_html(**kwargs)

    def __html__(self):
        return self.as_html()
