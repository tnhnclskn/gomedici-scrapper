import requests
from bs4 import BeautifulSoup
import re
import sys
import xlsxwriter


class GoMedici:
    def __init__(self, session, companyLinks=[], page=1):
        self.session = session
        self.companyLinks = companyLinks
        self.page = page
        self.companies = []
        self.companyId = 0

    def getCompanyId(self):
        self.companyId += 1
        return self.companyId

    def getHeader(self, auth=False):
        return {
            'cookie': '_medici_session=cf25acf9907d87f608b9893d69a622111;',
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'en-US,en;q=0.5',
            'connection': 'keep-alive',
            'cookie': '_medici_session=190d276b106cf3269e5b8915f92ac40e',
            'host': 'gomedici.com',
            'upgrade-insecure-requests': '1',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:81.0) Gecko/20100101 Firefox/81.0'
        }

    def getPage(self, pageNum):
        url = 'https://gomedici.com/companies?model=Companies&page=' + \
            str(pageNum) + '&size=9&term=*&type=Filters'
        response = requests.get(url, headers=self.getHeader(True))
        return response.text

    def parseCompanyLinks(self, content):
        soup = BeautifulSoup(content, 'html.parser')
        links = []
        for company in soup.find_all(class_="col-md-4"):
            links.append(company.select_one("#cp-img > a")['href'])
        return links

    def fetchCompanyLinks(self):
        while True:
            pageContent = self.getPage(self.page)
            pageLinks = self.parseCompanyLinks(pageContent)
            self.companyLinks += pageLinks
            print('Fetched Page Number:', self.page)
            self.page += 1
            if len(pageLinks) == 0:
                break

    def fetchCompaniesData(self):
        for companyLink in self.companyLinks:
            company = self.fetchCompanyData(companyLink)
            print('Fetched Company:', company['COMPANY_NAME'])
            self.companies.append(company)

    def fetchCompanyData(self, companyLink):
        url = 'https://gomedici.com' + companyLink
        response = requests.get(url, headers=self.getHeader(True))
        soup = BeautifulSoup(response.text, 'html.parser')

        categories = []
        for category in soup.select('p.company_sub_title_head > a'):
            categories.append(category.string)

        founders = []
        for founder in soup.select('#cp__data__people > *'):
            if founder.name == 'p':
                founders.append(founder.string)
            elif len(founders) > 0:
                break

        employeesRecord = False
        employees = ''
        for row in soup.select('#cp__data__people > *'):
            if employeesRecord:
                employees = row.string
            employeesRecord = row.string == 'NUMBER OF EMPLOYEES'

        totalFundingAmount = ''
        totalFundingAmountRecord = False
        for row in soup.select('#cp__data_momentum > *'):
            if totalFundingAmountRecord:
                totalFundingAmount = row.string
            totalFundingAmountRecord = row.string == 'TOTAL FUNDING'

        investors = []
        for investor in soup.select('#cp__data__people > p > a'):
            investors.append(investor.string)

        website = ''
        linkedin = ''
        facebook = ''
        twitter = ''
        for link in soup.select('ul.technologies > li > a'):
            href = link['href']
            linkType = link.select_one('img')['alt']
            if linkType == 'website':
                website = href
            elif linkType == 'linkedin':
                linkedin = href
            elif linkType == 'facebook':
                facebook = href
            elif linkType == 'twitter':
                twitter = href

        relateds = []
        for related in soup.select('.related_companies .font-weight-bold > a'):
            relateds.append(related.string)

        return {
            'ID': self.getCompanyId(),
            'COMPANY_NAME': soup.select_one('.company_title_head').string,
            'URL': url,
            'COMPANY_LOGO': soup.select_one('.cp-detail-content-img')['src'],
            'DESCRIPTION': soup.select_one('#cp__data_about > p').string,
            'PRODUCT': soup.select_one('#cp__data_focus > p').string,
            'LOCATION': soup.select_one('div.company_sub_title_head:nth-child(5) > span:nth-child(1) > strong:nth-child(1)').string,
            'HQ_COUNTRY': soup.select_one('div.company_sub_title_head:nth-child(5) > span:nth-child(1) > strong:nth-child(1)').string,
            'FOUNDED': soup.select_one('div.company_sub_title_head:nth-child(4) > span:nth-child(1) > strong:nth-child(1)').string,
            'CATEGORIES': categories,
            'FOUNDERS': founders,
            'WEBSITE_URL': website,
            'LINKEDIN_URL': linkedin,
            'FACEBOOK_URL': facebook,
            'TWITTER_URL': twitter,
            'EMPLOYEES': employees,
            'TOTAL_FUNDING_AMOUNT': totalFundingAmount,
            'INVESTORS': investors,
            'RELATED_COMPANIES': relateds,
            'SOURCE': url
        }

    def exportXls(self):
        workbook = xlsxwriter.Workbook('export/data.xlsx')
        worksheet = workbook.add_worksheet()

        # Some data we want to write to the worksheet.
        data = []
        data.append(self.companies[0].keys())
        for company in self.companies:
            values = []
            for value in company.values():
                if type(value) is list:
                    value = ';'.join(value)
                values.append(value)
            data.append(values)

        row = 0
        for line in data:
            col = 0
            for v in line:
                worksheet.write(row, col, v)
                col += 1
            row += 1
        workbook.close()


gomedici = GoMedici('190d276b106cf3269e5b8915f92ac40e')
gomedici.fetchCompanyLinks()
gomedici.fetchCompaniesData()
gomedici.exportXls()
