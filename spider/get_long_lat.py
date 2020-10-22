# -*- coding:utf-8 _*-

#   author : YOYO
#   time :  2020/10/21 11:20
#   email : youyou.xu@enesource.com
#   project_name :  get_city_jingweidu
#   file_name :  get_long_lat
#   function：查询企业经纬度并保存到excel

import requests
from doexcel.do_excel import DoExcel


class Spider:
    def main(self):
        excel = DoExcel()
        company_list = excel.read()
        n = 1
        for company in company_list:
            n += 1
            # headers = {'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'}
            url = "https://restapi.amap.com/v3/geocode/geo?key=902ae42170c02ad53fde5fc5ee23c246&address={}".format(company.company)
            resp = requests.request(method="GET", url=url)
            result = resp.json()
            if result['geocodes']:
                long_and_lat = result['geocodes'][0]["location"]
                longitude, latitude = long_and_lat.split(",", 1)
                district = result['geocodes'][0]['district']
                adcode = result['geocodes'][0]['adcode']
                excel.write(n, longitude, latitude, district, adcode)
            else:
                continue



if __name__ == '__main__':
    Spider().main()



