import requests
from bs4 import BeautifulSoup
import os
from googleapiclient.discovery import build
import xlsxwriter
from PIL import Image
import os

my_dir_for_img1="Directory for images goes here"
my_dir_for_img2="Directory for excel will go here"
API_KEY = "your google API key for v3"
youtube = build('youtube', 'v3', developerKey=API_KEY)
#excel sht
file_name = "Name.xlsx"
workbook = xlsxwriter.Workbook("{}{}".format(my_dir_for_img2,file_name))
worksheet = workbook.add_worksheet()
#the numbers (should) change according to how many songs you have on your playlist (i could make that aumatic but nah)
#column sizes
worksheet.set_column("A1:A20",10)
worksheet.set_column("B1:L20",50)
#row size
worksheet.set_default_row(68)

#90%chatgpt code that helped faster get youtube links cuz i had skill issue
def search_youtube(query):
    search_response = youtube.search().list(
        q=query,
        part='id,snippet',
        maxResults=1
    ).execute()

    video_link = []
    for item in search_response['items']:
        if item['id']['kind'] == 'youtube#video':
            video_link.append(f"https://www.youtube.com/watch?v={item['id']['videoId']}")


    return video_link
#Function that get the html data of any link ,10% from stack over flow
def get_html(spo_url):

 meFR = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }

 response = requests.get(spo_url, headers=meFR)
 response.raise_for_status()

 soup = BeautifulSoup(response.text,"html.parser")
 return(soup)


######## HERE GOES THE SPOTIFY PLAYLIST #########
soup = get_html("https://open.spotify.com/playlist/0Rr8J1Mxb9YGumNsSrpjNQ?si=bbe1239267b74ed4&nd=1")
a = soup.find_all("meta", attrs={"name": "music:song"})

url_img=[]
songs = []
Artists = []
link =[]

for i in a:
   k= i["content"] 
   s = get_html(k)
#gets the image
   soup3 = s.find("meta",attrs={"name":"twitter:image"})
   url_img.append(soup3["content"])
#song tittle
   soup1 = s.find("meta",attrs={"name":"twitter:title"})
   songs.append(soup1["content"])
#Artist
   soup2 = s.find("meta",attrs={"name":"twitter:description"})
   Artists.append(soup2["content"])

#It searches for the song of the artist on youtube and gets the first result
#its 90%~ accurate
   query = "{}{}{}".format(soup1["content"]," ",soup2["content"])
   video_links = search_youtube(query)
   link.append(video_links[0])

####insert data in excel#####
#first we download the images tho in our image directory
index = 0
albums_imgs = []
for i in url_img:
   Title = songs[index]
   Title = Title.replace(" ","")
   albums_imgs.append(Title)
   image_response = requests.get(i)
   index += 1
   with open("{}{}{}".format(my_dir_for_img1,Title,".jpg"),"wb")as file:
      file.write(image_response.content)

#insert every image in excel [had to resize them cuz it was weird]
row1 = 0
for album_cover in albums_imgs:
   IMAGE = "{}{}{}".format(my_dir_for_img1,album_cover,".jpg")
   size = Image.open(IMAGE)
   Re_size = size.resize((125,125))
   x = '{}{}{}{}'.format(my_dir_for_img1,albums_imgs[row1],"C",".jpg")
   Re_size.save(x)
   worksheet.insert_image(row1,0,x,{'x_scale':0.6,'y_scale':0.6,'y_offset':5,'positioning':1})
   row1 +=1
#row1 can work as an index too  (just realized that :skull emoji:)
row1 = 0
for Artist in Artists:
   worksheet.write(row1,1,Artist)
   worksheet.write(row1,2,songs[row1])
   worksheet.write(row1,3,link[row1])

   row1 +=1

#by closing it we save it too
workbook.close()

#delete the images since we dont nee them anymore (we can delete them every time we insert but agin im lazy)
Wedont_need_them = os.listdir(my_dir_for_img1)
for dir in Wedont_need_them:
    fullpath = "{}{}".format(my_dir_for_img1,dir)
    os.remove(fullpath)