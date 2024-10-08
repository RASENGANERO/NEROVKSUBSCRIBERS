import vk_api
import csv
import easygui
import time
import xlsxwriter
import os
import math
import random
from io import BytesIO
from urllib.request import urlopen
from PIL import Image


LOGIN_VK = 'www.223@list.ru'
PASSWORD_VK = 'Redfusi2178905609-45'

LKST_ERROR=list()
FIELDS='bdate,books,career,connections,contacts,education,exports,games,interests,movies,music,occupation,relation,about,activities,followers_count,'
PHOTOS='photo_200_orig,photo_200,photo_400_orig,photo_max,photo_max_orig'


def captcha_handler(captcha):
    s=str(captcha.get_url())
    f=open(r'captch.jpg',"wb")        
    cur=requests.get(s)
    f.write(cur.content) 
    f.close()
    os.startfile('captch.jpg')
    key = input("Введите код капчи : ")
    os.system("taskkill /IM dllhost.exe")
    os.remove('captch.jpg')
    return captcha.try_again(key)

def auth_handler():
    key = input("Введите код: ")
    remember_device = True 
    return key, remember_device


vk_session = vk_api.VkApi(LOGIN_VK, PASSWORD_VK,
    auth_handler=auth_handler,    # функция для обработки двухфакторной аутентификации
    captcha_handler=captcha_handler,  # функция для обработки капчи
    token='95fe5e8a95fe5e8a95fe5e8ab095877a85995fe95fe5e8af4d29c7aedcb50e830d6d594')
vk_session.auth()
vk=vk_session.get_api()


def get_all_members(group_id,vk):
    members = vk.groups.getMembers(group_id=group_id, fields='city,country,home_town,sex,'+FIELDS+PHOTOS)
    count = members['count']
    offset = 1000
    members = [members['items']]
    while offset < count:
        members.extend([vk.groups.getMembers(group_id=group_id, fields='city,country,home_town,sex,'+FIELDS+PHOTOS, count=1000, offset=offset)['items']])
        offset += 1000
    return members


def set_close(k):
    if k==False:
        return 'Открытая'
    else:
        return 'Закрытая'
    

def set_sex(k):
    if k==1:return 'Женский'
    if k==2:return 'Мужской'
    if k!=1 and k!=2: return 'Не укзано'


def set_sheet(k,glob_list):
    workbook = xlsxwriter.Workbook(k)#save_file)
    worksheet = workbook.add_worksheet('Информация о подписчиках')
    bold=workbook.add_format({'align':'center','valign':'vcenter', 'text_wrap':True, 'bold':True})
    worksheet.write(0, 0, 'Фото профиля',bold)
    worksheet.write(0, 1, 'ID',bold)
    worksheet.write(0, 2, 'Тип страницы',bold)
    worksheet.write(0, 3, 'Имя',bold)
    worksheet.write(0, 4, 'Фамилия',bold)
    worksheet.write(0, 5, 'Пол',bold)
    worksheet.write(0, 6, 'Страна',bold)
    worksheet.write(0, 7, 'Город',bold)
    worksheet.write(0, 8, 'СП',bold)
    worksheet.write(0, 9, 'Интересы',bold)
    worksheet.write(0, 10, 'Книги',bold)
    worksheet.write(0, 11, 'О себе',bold)
    worksheet.write(0, 12, 'Игры',bold)
    worksheet.write(0, 13, 'Фильмы',bold)
    worksheet.write(0, 14, 'Занятие',bold)
    worksheet.write(0, 15, 'Музыка',bold)
    worksheet.write(0, 16, 'Моб. тел.',bold)
    worksheet.write(0, 17, 'Дом. тел.',bold)
    worksheet.write(0, 18, 'Подписчики',bold)
    worksheet.write(0, 19, 'Карьера',bold)
    worksheet.write(0, 20, 'Университет',bold)
    worksheet.write(0, 21, 'Факультет',bold)
    worksheet.write(0, 22, 'Родной город',bold)
    worksheet.write(0, 23, 'Образование',bold)
    worksheet.write(0, 24, 'Facebook',bold)
    worksheet.write(0, 25, 'Instagram',bold)
    worksheet.write(0, 26, 'Twitter',bold)
    worksheet.write(0, 27, 'Skype',bold)
    worksheet.set_default_row(141)    
    worksheet.set_column('A:A',26)
    worksheet.set_column('B:H',30)
    bold = workbook.add_format({'align':'center','valign':'vcenter','text_wrap':True})

    
    for q in range(len(glob_list)):
        sk=glob_list[q]

        for p in range(len(sk)):
            if p==0:
                worksheet.insert_image(int(q+1),0,sk[p][0],sk[p][1])
            else:
                worksheet.write(int(q+1),p,sk[p],bold)
    workbook.close()



def check_of_none(s):
    lkst=['bdate','interests','books','about','games','movies','activities','music',
	'mobile_phone','home_phone','followers_count','university_name','faculty_name',
	'home_town','facebook','twitter','instagram','skype']

    for v in range(len(lkst)):
        if lkst[v] not in s:
            s.update({lkst[v]:"Неизвестно"})
        if lkst[v] in s and s[lkst[v]]=='':
            s[lkst[v]]="Неизвестно"
        
    

    s1=['city','country','occupation']
    s2=['title','title','name']
    for v in range(len(s1)):
        if s1[v] not in s:
            s.update({s1[v]:{s2[v]:"Неизвестно"}})


    if 'relation' not in s:
        s.update({'relation':0})

    if 'career' not in s:
        s.update({'career':[]})


    s1=['photo_50','photo_100','photo_200_orig','photo_200','photo_400_orig','photo_max','photo_max_orig']
    for v in range(len(s1)):
        if s1[v] not in s:
            s.update({s1[v]:"https://vk.com/images/camera_400"})
    return s

    
def clear_ban(s):
    for v in range(len(s)):
        #print(s[v])
        if 'deactivated' in s[v]:
            s[v]=None
        else:s[v]=s[v]
    return list(filter(None,s))


def get_photo(s):
    s=list(filter(None,s))
    for v in range(len(s)):
        if str(s[v]).startswith("https://vk.com/images/camera")==True:
            s[v]=None
    s=list(filter(None,s))
    if len(s)==0:
        return 0
    else:
        #print(s)
        try:
            s=str(random.choice(s))
        #s=str(random.choiсe()
        
            image_data = BytesIO(urlopen(s).read())
            im = Image.open(image_data,'r')
            max_size = (230, 189)
            im.thumbnail(max_size, Image.ANTIALIAS)
            imgByteArr = BytesIO()
            im.save(imgByteArr,format='PNG')
            return [s, {'image_data':imgByteArr}]
        except Exception:
            return 0

def set_all(s,s1,s2):
    for v in range(len(s1)):
        if s==s1[v]:s=s2[v]
    return s

def set_social(s,s1):
    if s!="Неизвестно":
        return str(s1)+str(s)
    else:
        return s

def set_career(s):
    if len(s)==0:
        return "Неизвестно"
    else:
        for v in range(len(s)):
            if 'company' in s[v]:
                s[v]=str(s[v]['company'])
            
            if 'group_id' in s[v]:
                s[v]='https://vk.com/club'+str(s[v]['group_id'])
        return ',\n'.join(s)

def get_id_group(s,vk):
    if s.startswith('https://vk.com/club')==True:
        return int(str(s).split('https://vk.com/club')[-1])
    else:
        return int(vk.groups.getById(group_id=str(str(s).split('https://vk.com/')[-1]))[0]['id'])

if __name__=='__main__':

    save_dir=str(easygui.diropenbox(title='Укажите папку для данных'))

    id=get_id_group(str(input('Введите ссылку на сообщество: ')),vk)
    lkstid=get_all_members(id,vk)
    


    print('Все ID страниц получены!')

    col=int(0)
    for v in range(len(lkstid)):
        user_data=lkstid[v]
        user_data=clear_ban(user_data)
        glob_list=list()
        col+=len(user_data)
        for q in range(len(user_data)):
            user_data[q]=check_of_none(user_data[q])
            photo_p=get_photo([user_data[q]['photo_200_orig'],
                               user_data[q]['photo_400_orig'],user_data[q]['photo_max'],user_data[q]['photo_max_orig']])

            if photo_p!=0:
                user_data[q]['id']='https://vk.com/id'+str(user_data[q]['id'])
                user_data[q]['is_closed']=set_all(user_data[q]['is_closed'],[False,True],['Открытая','Закрытая'])
                user_data[q]['sex']=set_all(user_data[q]['sex'],[1,2,0],["Женский","Мужской","Не указано"])
                user_data[q]['relation']=set_all(user_data[q]['relation'],[1,2,3,4,5,6,7,8,0],["не женат/не замужем",
                                                                              "есть друг/есть подруга",
                                                                              "помолвлен/помолвлена",
                                                                              "женат/замужем",
                                                                              "всё сложно",
                                                                              "в активном поиске",
                                                                              "влюблён/влюблена",
                                                                              "в гражданском браке",
                                                                              "не указано"])
                user_data[q]['career']=set_career(user_data[q]['career'])
                user_data[q]['facebook']=set_social(user_data[q]['facebook'],"https://www.facebook.com/")
                user_data[q]['instagram']=set_social(user_data[q]['instagram'],"https://www.instagram.com/")
                user_data[q]['twitter']=set_social(user_data[q]['twitter'],"https://twitter.com/")
                     
                lkj=[photo_p,user_data[q]['id'],user_data[q]['is_closed'],user_data[q]['first_name'],
                     user_data[q]['last_name'],user_data[q]['sex'],user_data[q]['country']['title'],
                     user_data[q]['city']['title'],user_data[q]['relation'],user_data[q]['interests'],
		             user_data[q]['books'],user_data[q]['about'],user_data[q]['games'],user_data[q]['movies'],
		             user_data[q]['activities'],user_data[q]['music'],user_data[q]['mobile_phone'],
                     user_data[q]['home_phone'],user_data[q]['followers_count'],user_data[q]['career'],
		             user_data[q]['university_name'],user_data[q]['faculty_name'],user_data[q]['home_town'],
                     user_data[q]['occupation']['name'],user_data[q]['facebook'],user_data[q]['instagram'],
                     user_data[q]['twitter'],user_data[q]['skype']]
                
                glob_list.append(lkj)  
                print('Обработано подписчиков: ',q,' из ',len(user_data),' всего: ',col,' часть ',v, ' из ',len(lkstid))
        set_sheet(save_dir+'\\'+str(v+1)+'.xlsx',glob_list)