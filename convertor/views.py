from django.shortcuts import render
from .forms import Profile_Form
from .models import User_Profile
IMAGE_FILE_TYPES = ['png', 'jpg', 'jpeg', 'pdf', 'docx']
import pandas as pd
import numpy as np
import re
import zipfile
from docx.shared import Pt
from docx import Document
from docx2python import docx2python
from docx2csv import extract_tables, extract
from docx2python.iterators import iter_paragraphs
import os.path
import base64
import dropbox
import convertapi


def create_profile(request):
    form = Profile_Form()
    if request.method == 'POST':
        form = Profile_Form(request.POST, request.FILES)
        if form.is_valid():
            #form.save()
            user_pr = form.save(commit = False)
            user_pr.display_picture = request.FILES['display_picture']
            file_type = user_pr.display_picture.url.split('.')[-1]
            file_type = file_type.lower()
            if file_type not in IMAGE_FILE_TYPES:
                return render(request, 'convertor/error.html')
            else:
                user_pr.save()
                '''pathy = user_pr.display_picture.url

                fp = open( pathy, 'rb')        
                data = fp.read()
                variables = RequestContext(request, {'file': data})
                output = template.render(variables)                
                return HttpResponse(output)'''

                '''we have the processing code out here'''

                docx = 'media/'+ str(user_pr.display_picture)
                print(docx)
                filename,extension = os.path.splitext(docx)
                x=np.random.randint(0,9999999,1)
                x=str(x[0])
                access_token = 'enter your token'
                fbt= access_token.encode('ascii')
                mbt = base64.b64decode(fbt)
                at = mbt.decode('ascii')
                dbx = dropbox.Dropbox(at)
                with open(filename+"_key.csv","wb")  as f:
                    metadata,res = dbx.files_download("/key.csv")                
                    f.write(res.content)

                qt=0
                key_data=pd.read_csv(filename+"_key.csv")
                key=str(key_data.iloc[0,0])

                if extension!='.docx' and extension!='.doc':
                    
                    convertapi.api_secret = key
                    result = convertapi.convert('docx', { 'File': docx})
                    result.file.save(filename+'.docx')
                    docx=filename+'.docx'
                    qt=2


                img_size=0
                m=[]
                text=0
                tb=0
                txt_line=0
                name_p=[]
                fonttxt_line=[]
                ln_id=[]
                mail=[]



                #no of images


                doc = docx2python(docx)
                img_size=len(doc.images)
                print('img',img_size)

                #mob

                d1=[]
                for i in range(len(doc.body)):
                    d1.extend(doc.body[i])

                d2=[]
                for i in range(len(d1)):
                    d2.extend(d1[i])
                d3=[]
                for i in range(len(d2)):
                    d3.extend(d2[i])

                regex=  '(\(?\d{3}\D{0,3}\d{3}\D{0,3}\d{4}).*?'
                regex1="\w\d \w\w \w\w \w\w \w\d|(?<=[^\d][^_][^_] )[^_]\d[^ ]\d[^ ][^ ]+|(?<= [^<]\w\w \w\w[^:]\w[^_][^ ][^,][^_] )(?: *[^<]\d+)+"

                for i in d3:
                   
                    if re.search(regex,i):
                        
                        if re.search(regex1, i):
                              m=re.search(regex1, i)
                              m=m.group()
                        else:
                            m=re.search(regex,i)
                            m=m.group()
                            
                        break
                print('mob',m)




                #no of text
                text=[]
                for i in d3:
                    result = re.findall(r"[a-zA-z]+",i )
                    text.extend(result)
                text=len(text)
                if img_size>0:
                    text=text-2*img_size
                print('text',text)



                # no of tables

                tables = extract_tables(docx)
                tb=0
                for table in tables:
                    if len(table)>2:
                       
                       for cell in table:
                            y=[]
                            if len(cell)>2:
                               for i in range(len(cell)-1):
                                   if cell[i]==cell[i+1]:
                                        y.append("col")
                                        break
                            else:
                               break
                      
                       if len(y)>0 and len(y)<=len(cell)/2:
                           tb=tb+1
                               
                print('table',tb)    

                #no of line


                l=[]
                k=[]
                regex="\w*[A-Z]\w*[A-Z]\w*"
                for st in d3:
                    s=0
                    r = re.sub( "[^a-zA-Z]" ,' ',st)
                    r=r.split()
                    r=' '.join(r)
                   
                    if re.search(regex,r):
                       s=re.findall(regex,r)
                       s=' '.join(s)
                       

                    if len(str(s))>15:
                         continue
                    
                    k.append(len(r))
                    l.append(r)    
                    
                    
                for j,i in enumerate (k):
                    if i>15:
                        txt_line=txt_line+1
                       
                print('line',txt_line)




                #name of person
                z=1
                regex='[A-Za-z]{2,25}\s[A-Za-z]{2,25}'
                for i,st in enumerate(d3):
                    if re.search(regex,st):
                         name_p=re.findall(regex,st)
                         break
                    if i==5:
                       z=2
                       break
                if len(name_p)>1 or z==2:
                    d4=d3[::-1]
                    for i,st in enumerate(d4):
                        if re.search(regex,st):
                             name_p=re.findall(regex,st)
                        if len(name_p)==1:
                                break
                        if i==5:
                            break
                print('name',name_p)




                #font shape and size

                fonts=[]
                t=0
                document = Document(docx)
                for p in document.paragraphs:
                    #print(p)
                    s=0
                    name = p.style.font.name
                    size = p.style.font.size
                   
                    for i,font in enumerate(fonts):
                         if font[0]==name:
                             s=s+1
                        
                    if s==0:
                        if name==None and t==0:
                            name='Arial'
                            t=t+1
                        if name==None and t>0:
                            continue
                        if size==None:
                            size=Pt(11)
                            
                        if size!=None:
                            size=size.pt
                        fonts.append([name,size])

                fonts=dict(fonts)       
                print('font',fonts)




                #linkldin
                document=zipfile.ZipFile(docx)
                xml_content = document.read('word/document.xml')
                xml_str = str(xml_content)
                link_list = re.findall('http.*?\<',xml_str)[1:]
                link_list = [x[:-1] for x in link_list]

                p = re.compile('((http(s?)://)*([www])*\.|[linkedin])[linkedin/~\-]+\.[a-zA-Z0-9/~\-_,&=\?\.;]+[^\.,\s<]')

                for i in link_list:
                    if p.match(i)!=None:
                        ln_id=i
                        break    
                    

                regex='(http(s?)://|[a-zA-Z0-9\-]+\.|[linkedin])[linkedin/~\-]+\.[a-zA-Z0-9/~\-_,&=\?\.;]+[^\.,\s<]'

                if len(ln_id)==0:
                    
                    for i,st in enumerate(d3):
                        st=st.lower()
                        if re.search(regex,st):
                                
                            ln_id=[x.group() for x in re.finditer( regex,st)]
                           
                    s=0
                    for i in range(len(ln_id)):
                        if len(ln_id[i])>s:
                            s=len(ln_id[i])
                            z=i
                    if len(ln_id)==0:
                        pass
                    elif len(ln_id[z])>15:
                        ln_id=ln_id[z] 
                    else:
                         ln_id=[]

                print('ln',ln_id)

                #mail
                document.close()
                regex='[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
                mail_id = re.findall(regex,xml_str)
                mail= [x[:] for x in mail_id]

                if len(mail)==0:
                     for st in d3:
                         if re.search(regex,st):
                            mail=re.search(regex,st)
                            mail=mail.group()
                            break

                print('mail',mail)


                if len(name_p)==0:
                    name_p=None

                if len(m)==0:
                    m=None

                if len(name_p)==0:
                   name=None

                if len( fonts)==0:
                   fonts=None

                if len(ln_id)==0:
                   ln_id=None

                if len(mail)==0:
                   mail=[None]


                result=pd.DataFrame(columns=['Name','linkedln_id','mail_id','mob','no_of_img','fonts_name_and size','no_of_text_line','no_of_text','no_table'])

                result.loc['result',:]=name_p[0],ln_id,mail[0],m,img_size,fonts,s,text,tb


                #save file
                file = result.to_csv('media/result.csv',index=False)


                path='media/result.csv'


                class TransferData:
                    def __init__(self, at):
                        self.at = at

                    def upload_file(self, file_from, file_to,file_from_r,file_to_r):
                        
                        dbx_1 = dropbox.Dropbox(self.at)
                        dbx_2 = dropbox.Dropbox(self.at)

                        with open(file_from, 'rb') as f:
                               dbx_1.files_upload(f.read(), file_to)
                        with open(file_from_r, 'rb') as f:
                               dbx_2.files_upload(f.read(), file_to_r)    
                    def upload_file_(self,file_from_w,file_to_w):
                        
                        dbx = dropbox.Dropbox(self.at)

                        with open(file_from_w, 'rb') as f:
                               dbx.files_upload(f.read(), file_to_w)       

                    
                def pdf():
                    transferData = TransferData(at)
                    file_from_w=docx
                    file_to_w='/'+x+'converted.docx'
                   
                    transferData.upload_file_(file_from_w,file_to_w)


                def main():
                    transferData = TransferData(at)
                    
                    file_from =path
                    file_to = ('/'+x+'result.csv' )
                    file_from_r=docx
                    file_to_r=('/'+x+'resume'+extension)
                   
                    transferData.upload_file(file_from, file_to,file_from_r,file_to_r)
                    if qt==2:
                        pdf()


                main()






                print(result)
                #file = result.to_csv('media/result.csv',index=False)
                #file = file + '/result.csv'
                
















                return render(request, 'convertor/details.html', {'user_pr': user_pr,})
            
            


            
    context = {"form": form,}
    return render(request, 'convertor/create.html', context)
