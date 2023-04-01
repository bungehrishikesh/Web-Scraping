from bs4 import BeautifulSoup
import requests, openpyxl

#2nd step - loading data into excel

#To create a new excel file
excel = openpyxl.Workbook()
print(excel.sheetnames)

#To make sure we are working on the active sheet as there can be more than 1 sheets
sheet = excel.active

sheet.title = 'Top Rated Movies'
print(excel.sheetnames)

#Column headers
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])

# 1st Step - fetching data
try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status() 

    soup = BeautifulSoup(source.text,'html.parser')
    
    #TO access the tbody tag as all the movies are inside tbody tag and then usinf find_all to find tr tag
    movies = soup.find('tbody', class_="lister-list").find_all('tr') 

    #print(len(movies))   //250 as there are 250 top rated movies 
    
    for movie in movies:
        
        #<td class="titleColumn">
        #1.
        #<a href="/title/tt0111161/?pf_rd_m=A2FGELUUNOQJNL&amp;pf_rd_p=1a264172-ae11-42e4-8ef7-7fed1973bb8f&amp;pf_rd_r=V1GQWSSHK95MQE3NNSNX&amp;pf_rd_s=center-1&amp;pf_rd_t=15506&amp;pf_rd_i=top&amp;ref_=chttp_tt_1" title="Frank Darabont (dir.), Tim Robbins, Morgan Freeman">The Shawshank Redemption</a>
        #    <span class="secondaryInfo">(1994)</span>
        #</td>
        name = movie.find('td', class_="titleColumn").a.text
        #strip will strip all the new line characters #1.Shawshank Redemption (1994) split=>['1','Shawshank Redemption'] [0]will fetch only 1st index i.e 1
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0] 

        year = movie.find('td', class_="titleColumn").span.text.strip('()')

        #<td class="ratingColumn imdbRating">
            #<strong title="9.2 based on 2,720,957 user ratings">9.2</strong>
        #</td>
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

#last step
excel.save('IMDB Movie Ratings.xlsx')