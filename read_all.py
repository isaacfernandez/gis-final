import geoip2.database
import xlsxwriter
import time

class Loc(object):
    """docstring for Loc"""
    def __init__(self, ip, country, co_code, city, lat, lon):
        super(Loc, self).__init__()
        self.ip = ip
        self.country = country
        self.co_code = co_code
        self.city = city
        self.lat = lat
        self.lon = lon

def add_worksheet(wb, file_handle, title):
    ws = workbook.add_worksheet(title)
    input_file = open(file_handle, 'r').read().splitlines()
    write_worksheet(ws, input_file)

#consumes an OpenBL list excel workbook handler and inserts appropiate worksheets
def write_worksheet(ws, record):
    locations = []
    record = record[4:]
    for ip in record:
        try:
            res = reader.city(ip)
            country = res.country.name
            co_code = res.country.iso_code
            city = res.city.name
            lat = res.location.latitude
            lon = res.location.longitude
            locations.append(Loc(ip, country, co_code, city, lat, lon))
        except:
            next
            #YOLO
    set_headers(ws)
    row = 1
    for location in locations:
        write_loc(ws, location, row)
        row += 1

def write_loc(ws, loc, row):
    col = 0
    ws.write(row, col,   loc.ip)
    ws.write(row, col + 1,   loc.country)
    ws.write(row, col + 2,   loc.co_code)
    ws.write(row, col + 3,   loc.city)
    ws.write(row, col + 4,   loc.lat)   
    ws.write(row, col + 5,   loc.lon)  

def set_headers(ws):
    row = 0
    col = 0
    ws.write(row, col, "ip address")
    ws.write(row, col + 1, "country")
    ws.write(row, col + 2, "country code")
    ws.write(row, col + 3,  "city")
    ws.write(row, col + 4, "lat")
    ws.write(row, col + 5, "lon")
    ws.write(row, col + 6, "time")

t0 = time.time()
reader =  geoip2.database.Reader('db/GeoLite2-City.mmdb')
print("Beginning process of all")
workbook = xlsxwriter.Workbook('openBL_geospatial.xlsx')
print("Workbook created")
print("Beginning all file")
add_worksheet(workbook, 'db/base_all.txt', 'all')
print("Beginning ssh-only file")
add_worksheet(workbook, 'db/base_all_ssh-only.txt', 'ssh')
print("Beginning ftp-only file")
add_worksheet(workbook, 'db/base_all_ftp-only.txt', 'ftp')
print("done, closing wb")
workbook.close()
t1 = time.time()
total = t1 - t0
print("Runs over. Duration: %d", total)
#f = open('db/base_all.txt', 'r').read().splitlines()
#for ip in f:
 #   res = reader.city(ip)
  ######ations.append(Loc(ip, country, co_code, city, lat, lon))