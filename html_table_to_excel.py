__author__ = 'root'

from django.http import HttpResponse
from django.utils.http import urlquote
import xlwt, HTMLParser, StringIO, uuid

blue_stype = xlwt.easyxf('alignment: horz left, vert top; pattern: pattern solid, fore_colour light_blue; font: bold on;')
red_stype = xlwt.easyxf('alignment: horz left, vert top; pattern: pattern solid, fore_colour red; font: bold on;')
green_stype = xlwt.easyxf('alignment: horz left, vert top; pattern: pattern solid, fore_colour light_green; font: bold on;')
yellow_stype = xlwt.easyxf('alignment: horz left, vert top; pattern: pattern solid, fore_colour light_yellow; font: bold on;')
orange_stype = xlwt.easyxf('alignment: horz left, vert top; pattern: pattern solid, fore_colour light_orange; font: bold on;')
merge_stype = xlwt.easyxf('alignment: wrap on;')

bold_stype = xlwt.easyxf('alignment: horz left, vert top; font: bold on;')
mydefault_stype = xlwt.easyxf('alignment: horz left, vert top')

STAG = 'stag'
ETAG = 'etag'
DATA = 'data'

def html_table_to_excel(table):
    """ html_table_to_excel(table): Takes an HTML table of data and formats it so that it can be inserted into an Excel Spreadsheet.
    """
    table_ls = []

    class MyHTMLParser(HTMLParser.HTMLParser):
        '''
        parser
        '''

        def handle_starttag(self, tag, attrs):
            table_ls.append((STAG, tag, attrs))

        def handle_endtag(self, tag):
            table_ls.append((ETAG, tag, None))

        def handle_data(self, contentstr):
            table_ls.append((DATA, contentstr.strip(), None))

    p = MyHTMLParser()
    p.feed(table)

    return table_ls


def export_to_sheet(wb, sheet_title, table_str):
    '''
    sheet
    '''
    ws = wb.add_sheet(sheet_title)

    ls = html_table_to_excel(table_str)

    xstatus = ''
    cline = 0
    ccell = 0
    b_readyinsert = False
    xattrs = None
    xcontent = ''
    cells_occupy = set()

    for tag, content, attrs in ls:
        if tag == STAG and content == 'thead':
            xstatus = 'thead'
        elif tag == ETAG and content == 'thead':
            xstatus = ''
        if tag == STAG and content == 'tbody':
            xstatus = 'tbody'
        elif tag == ETAG and content == 'tbody':
            xstatus = ''

        elif tag == STAG and content == 'tr':
            # row go
            ccell = 0
        elif tag == ETAG and content == 'tr':
            # row go
            cline += 1

        elif tag == STAG and content in ['td', 'th']:
            b_readyinsert = True
            xcontent = ''
            xattrs = dict(attrs)
        elif tag == ETAG and content in ['td', 'th']:
            # fill cell
            xstyle = mydefault_stype
            if 'class' in xattrs:
                xattr_class = xattrs['class']
                if 'success' in xattr_class:
                    xstyle = green_stype
                elif 'warning' in xattr_class:
                    xstyle = yellow_stype
                elif 'danger' in xattr_class:
                    xstyle = red_stype
                elif 'info' in xattr_class:
                    xstyle = blue_stype
            if not xstyle:
                xstyle = xstatus == 'thead' and bold_stype or mydefault_stype

            # test occupy
            while (cline, ccell) in cells_occupy:
                ccell += 1

            if 'colspan' in xattrs and 'rowspan' in xattrs:
                rowspan = int(xattrs['rowspan'])
                colspan = int(xattrs['colspan'])
                ws.write_merge(cline, rowspan - 1 + cline, ccell, colspan - 1 + ccell, xcontent, xstyle)
                for x in range(0, rowspan):
                    cells_occupy.add((cline + x, ccell))
                for x in range(0, colspan):
                    cells_occupy.add((cline, ccell + x))
            elif 'rowspan' in xattrs:
                rowspan = int(xattrs['rowspan'])
                ws.write_merge(cline, rowspan - 1 + cline, ccell, ccell, xcontent, xstyle)
                for x in range(0, rowspan):
                    cells_occupy.add((cline + x, ccell))
            elif 'colspan' in xattrs:
                colspan = int(xattrs['colspan'])
                ws.write_merge(cline, cline, ccell, colspan - 1 + ccell, xcontent, xstyle)
                for x in range(0, colspan):
                    cells_occupy.add((cline, ccell + x))
            else:
                ws.write(cline, ccell, xcontent, xstyle)
                cells_occupy.add((cline, ccell))

            b_readyinsert = False
            xattrs = {}
            xstyle = None
            # cell go
            # ccell += 1

        elif b_readyinsert:
            if content == 'br':
                if xcontent:
                    xcontent += '\r'

            elif tag == DATA:
                content = content.strip()
                if content:
                    if xcontent:
                        xcontent += ' ' + content
                    else:
                        xcontent = content

    return wb


def export_to_xls(table, b_export_response=True, table_title=''):
    """
    @param  table:              string or dict
    """
    wb = xlwt.Workbook(style_compression=2)
    if isinstance(table, dict):
        if table:
            vx = ''
            for kt, v in table.items():
                vx += u'<table><thead><tr><th></th></tr><tr><th>{}</th></tr></thead></table>{}'.format(unicode(kt), v)
            export_to_sheet(wb, 'NEW', vx)
        else:
            wb.add_sheet('EMPTY')
    elif isinstance(table, list):
        if table:
            vx = ''
            for kt, v in table:
                vx += u'<table><thead><tr><th></th></tr><tr><th>{}</th></tr></thead></table>{}'.format(unicode(kt), v)
            export_to_sheet(wb, 'NEW', vx)
        else:
            wb.add_sheet('EMPTY')
    else:
        export_to_sheet(wb, 'NEW', table)

    if b_export_response:
        sio = StringIO.StringIO()
        wb.save(sio)
        dd = sio.getvalue()
        sio.close()

        #download
        response = HttpResponse(dd, content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename={0}.xls'.format(urlquote(table_title) or str(uuid.uuid4()))
        return response

    else:
        return wb