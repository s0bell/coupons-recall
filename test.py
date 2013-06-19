#!/usr/local/bin/python2.7
import json
import numpy
import sys
import urllib2

from time import strftime, sleep

from xlutils.copy import copy
from xlrd import open_workbook
from xlwt import easyxf


class CouponsRecallTest(object):

    def __init__(self):
        with open('config.json') as f:
            self.conf = json.load(f)
        self.workbook = open_workbook(self.conf['input_file'])
        self.search_retry_int = 30
        self.timestamp = strftime('%Y-%m-%dT%H-%M-%S')
        self.api_coupon_count = list()
        self.clipper_coupon_count = list()
        self.samples = 0

    def coupon_id_from_image_url(self, url):
        url_parts = url.split('/')
        return int(url_parts[-1].rstrip('.gif'))

    def coupon_index(self, sheet, row_index):
        coupons = dict()
        col_indices = [sheet.row_values(0).index(item) for item in
                       sheet.row_values(0) if item.lower().startswith('coupon')]
        for i in col_indices:
            if not sheet.cell(row_index, i).value:
                continue
            coupons[int(sheet.cell(row_index, i).value)] = i
        return coupons

    def print_stats(self):
        recall = format(100 * float(sum(self.api_coupon_count)) / float(
            sum(self.clipper_coupon_count)), '.2f')
        print '{0:16} {1}'.format('samples:', self.samples)
        print '{0:16} {1}'.format('api coupons:', sum(self.api_coupon_count))
        print '{0:16} {1}'.format('clipper coupons:', sum(self.clipper_coupon_count))
        print '{0:16} {1}%'.format('recall:', recall)
        print
        d = {'Relevant API coupons per item': self.api_coupon_count,
             'Relevant Clipper coupons per item': self.clipper_coupon_count}
        for key, val in d.items():
            print '-' * len(key), '\n', key, '\n', '-' * len(key)
            print '{0:7} {1}'.format('min:', int(min(val)))
            print '{0:7} {1}'.format('max:', int(max(val)))
            print '{0:7} {1}'.format('median:', int(numpy.median(val)))
            print
            for r in range(min(val), max(val) + 1):
                if val.count(r) == 0:
                    continue
                pct = format(100 * float(val.count(r)) / float(
                    self.samples), '.2f')
                print '{0:2} => {1}/{2} ({3}%)'.format(
                    r, val.count(r), self.samples, pct)
            print

    def query_api(self, upc):
        url = self.conf['coupons_api'] + str(upc)
        req = urllib2.Request(url)
        while True:
            try:
                response = urllib2.urlopen(req)
                return json.loads(response.read())
            except urllib2.URLError, e:
                print >> sys.stderr, e, url
                print >> sys.stderr, 'Retry in {0} seconds...'.format(
                    self.search_retry_int)
                sleep(self.search_retry_int)
                continue
            break

    def main(self):
        green = easyxf('pattern: pattern solid, fore-colour light_green')
        tmp_workbook = copy(self.workbook)
        r_sheet = self.workbook.sheet_by_index(0)
        w_sheet = tmp_workbook.get_sheet(0)
        col_upc = 0
        col_api_coupons = 3
        col_clipper_coupons = 4

        for row_index in range(1, r_sheet.nrows):
            if not r_sheet.cell(row_index, col_upc).value:
                continue  # Skip empty cells

            self.samples += 1
            upc = int(r_sheet.cell(row_index, col_upc).value)
            coupons = self.coupon_index(r_sheet, row_index)

            r = self.query_api(upc)

            matches = 0
            for coupon in r['coupons']:
                coupon_id = self.coupon_id_from_image_url(coupon['imageUrl'])
                if coupon_id in coupons:
                    matches += 1
                    w_sheet.write(row_index, coupons[coupon_id], coupon_id, green)

            w_sheet.write(row_index, col_api_coupons, matches)
            w_sheet.write(row_index, col_clipper_coupons,
                          len(coupons))

            self.clipper_coupon_count.append(len(coupons))
            self.api_coupon_count.append(matches)

        self.print_stats()

        tmp_workbook.save(self.conf['output_file'])


if __name__ == '__main__':
    CT = CouponsRecallTest()
    CT.main()
