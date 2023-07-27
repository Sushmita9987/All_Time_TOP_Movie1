from bs4 import BeautifulSoup
import requests,openpyxl

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie_name','Movie_Year','Movie_Rating'])

try:
    headers={'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'}
    url ="https://www.imdb.com/chart/top/"
    source = requests.get(url,headers= headers)
    source.raise_for_status()
    soup = BeautifulSoup(source.text,'html.parser')
    movies = soup.find('ul',class_ = "ipc-metadata-list ipc-metadata-list--dividers-between sc-3a353071-0 wTPeg compact-list-view ipc-metadata-list--base").find_all('li',class_="ipc-metadata-list-summary-item sc-bca49391-0 eypSaE cli-parent")
    for movie in movies:
        name = movie.find('a',class_ ="ipc-title-link-wrapper").h3.text
        # rank = movie.find('a',class_ ="ipc-title-link-wrapper").get_text(strip = True).split('.')[0]
        rating = movie.find('div',class_="sc-951b09b2-0 hDQwjv sc-14dd939d-2 fKPTOp cli-ratings-container").span.text
        year = movie.find('div',class_="sc-14dd939d-5 cPiUKY cli-title-metadata").span.text
        print(name,year,rating)
        sheet.append([name,year,rating]) 

except Exception as e:
    print(e)        
    
excel.save('IMDB Movie Rating.xlsx')    