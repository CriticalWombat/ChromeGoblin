import sqlite3
import xlsxwriter

class ChromeGoblin:
    def __init__(self):
        self.WB=xlsxwriter.Workbook('ChromeInfo.xlsx')
        self.WS=self.WB.add_worksheet(name='Google Searches')
        self.url_WS=self.WB.add_worksheet(name='Sites Visited')
        self.WS.write('A1', 'Google Searches')
        self.url_WS.write('A1', 'Sites Visited')
        self.row=0
        self.col=0
        s=self.get_searches()
        u=self.get_urls()
        self.__excelOps(s, u)
        self.WB.close()
        
    def get_urls(self):
        con = sqlite3.connect("History")
        cur = con.cursor()
        urls = []

        for row in cur.execute("SELECT * FROM urls"):
            url = str(row).split(",")[1]
            urls.append(url)
        return urls

    def get_searches(self):
        con = sqlite3.connect("History")
        cur = con.cursor()
        searches = []

        for row in cur.execute("SELECT * FROM keyword_search_terms"):
            search = str(row).split(",")[2]
            searches.append(search)
        return searches

    def __excelOps(self, searches, urls):
        for search in searches:
            self.WS.write(self.row +1, self.col, search)
            self.row +=1
        self.row = 0
        for url in urls:
            self.url_WS.write(self.row +1, self.col, url)
            self.row +=1

if __name__ == '__main__':
    ChromeGoblin()