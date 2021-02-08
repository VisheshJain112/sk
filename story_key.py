from os import curdir
from tkinter import *
import os
import tkinter.filedialog
import numpy as np
from styleframe import StyleFrame, utils
import math
import pandas as pd
import numpy as np

import csv
import matplotlib
from tkmagicgrid import *
from PIL import ImageTk, Image
from tkinter import ttk
import uuid
from itertools import count
from tkinter.tix import ScrolledWindow
import json
import requests
from styleframe import StyleFrame, utils
import random
import pandas as pd
import numpy as np
from sentence_splitter import SentenceSplitter, split_text_into_sentences
from gingerit.gingerit import GingerIt
from os import curdir
import textwrap
import re
import collections
import inflect
import tkinter.font as tkFont
# Designing window for registration



    

 
# Deleting popups
class ImageLabel(Label):
    """a label that displays images, and plays them if they are gifs"""
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        self.loc = 0
        self.frames = []

        try:
            for i in count(1):
                self.frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass

        try:
            self.delay = im.info['duration']
        except:
            self.delay = 100

        if len(self.frames) == 1:
            self.config(image=self.frames[0],borderwidth=0,highlightthickness = 0,pady=0,padx=0)
        else:
            self.next_frame()

    def unload(self):
        self.config(image="")
        self.frames = None

    def next_frame(self):
        if self.frames:
            self.loc += 1
            self.loc %= len(self.frames)
            self.config(image=self.frames[self.loc],borderwidth=0,highlightthickness = 0,pady=0,padx=0)
            self.after(self.delay, self.next_frame)

def delete_login_success():

    scan_screen.destroy()

    app = application_window()
    app.execute()


 
 
 
# Designing Main(first) window

class application_window():

    def __init__(self):
        self.name = "Mr. Smith"

        self.age = 32
        self.accompanied_by = "Mr. John"
        self.location = "Brighton High"
        self.category_logic = "Category-1,Category-2"
        self.struct_dict = {
                          "address_logic" : self.get_address_logic,
                          "report_logic" : self.get_report_logic,
                          "data" : self.get_data_logic,
                          "name" : self.get_name,
                          "age" : self.get_age,
                          "accompanied_by" : self.get_accompanied_by,
                          "location_data" : self.get_location,
                          "gender_pron" : self.get_pronoun,
                          "gender_det" : self.get_determiner,
                          "category_logic" : self.get_category_logic,
                          "also-1" : self.get_also1_logic,
                          "also-2" : self.get_also2_logic,
                          "category-1" : self.get_category_1,
                          "category-2" : self.get_category_2,
                          "category-3" : self.get_category_3,
                          "category-4" : self.get_category_4,
                          "category-5" : self.get_category_5,
                          "category-6" : self.get_category_6,
                          "category-7" : self.get_category_7,
                          "category-8" : self.get_category_8,
                          "category-9" : self.get_category_9
                          }

        self.idx = 0

        self.structure_logic_sheet = curdir + "/structure_logic_sheet.xlsx"
    def display_user_options_window(self):
        global img
        root = Tk()
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        root.geometry("800x800+%d+%d" % (screen_width/2-275, screen_height/2-125))
        root.configure(background='white')
        root.attributes('-topmost', True)
        #root.lift()
        #root.attributes("-fullscreen", True)  
        #root.configure(bg='white')
        ico_path = curdir+"\myicon.ico"
        root.iconbitmap(ico_path)

     
        root.title("Scanning AI")
        bg_path = curdir + "/scan.gif"
        lbl = ImageLabel(root)
        lbl.pack()
        lbl.load(bg_path)
        def read_path_info():

            file = open("google_path_info.txt","r")
            for lines in file.read().splitlines():
                if lines[0] == "T":
                    self.test_sheet = lines[1:]
                else:
                    self.story_logic_sheet = lines[1:]
            file.close() 
        ttk.Label(root, text = "Scanning the documents, Please wait.",  
        font = ("Times New Roman", 15)).pack()
        read_path_info()

        def display_output(text):
        
            root = Tk()
            root.geometry("1000x1000")
            root.title("Interpreted Story")
            
            curdir = os.getcwd()
            ico_path = curdir+"\myicon.ico"
            root.iconbitmap(ico_path)

            textExample=Text(root, height=40)
            textExample.pack()

            fontExample = tkFont.Font(family="Arial", size=16)

            textExample.configure(font=fontExample)
            textExample.insert(END,text)


        def del_and_open():
            self.test_sent = ""
            
            print("Succesfully Scanned")
            self.test_sheet = self.test_sheet.rstrip()
            self.test_sheet = self.test_sheet.lstrip()
          

            self.story_logic_sheet = self.story_logic_sheet.rstrip()
            self.story_logic_sheet = self.story_logic_sheet.lstrip()

            df_attempt = self.get_attempt_sheet()
            self.df_attempt = self.pad_array(df_attempt,self.row_lim+1)
            self.get_total_df()
            self.fetch_data()
            self.get_sequence_data()
            self.seq_attempt()

            clb_sheet = self.get_clubbing_sheet()
            self.clb_sheet = self.pad_array(clb_sheet,self.row_lim+1)
            remark_sheet = self.get_remark_sheet()
            self.remark_sheet = self.pad_array(remark_sheet,self.row_lim+1)
            dependency_sheet = self.get_dependency_sheet()
            self.dependency_sheet = self.pad_array(dependency_sheet,self.row_lim+1)
            category_sheet = self.get_category_sheet()
            self.category_sheet = self.pad_array(category_sheet,self.row_lim+1)
            data_logic_sheet = self.get_data_logic_sheet()
            self.data_logic_sheet = self.pad_array(data_logic_sheet,self.row_lim+1)
            self.get_os()


            total_sent = self.os+ "." +"\n"+ " "+self.get_data_dict()
            total_sent = total_sent.replace(' .','')
            total_sent = total_sent.replace("\n\n. ","\n\n")
            total_sent = self.final_processing(total_sent)
            total_sent = textwrap.dedent(total_sent)




            display_output(total_sent)

            

        def get_case_num():
            self.case_screen = Tk()

            self.case_screen.title("Case number")
            self.case_screen.geometry("300x250")
            curdir = os.getcwd()
            ico_path = curdir+"\myicon.ico"
            self.case_screen.iconbitmap(ico_path)
            Label(self.case_screen, text="Please enter Case Number below").pack()
            Label(self.case_screen, text="").pack()

      
            textBox1=Text(self.case_screen, height=2, width=20)
            
            textBox1.pack()
            def retrieve_input():
   
                self.case_num=textBox1.get("1.0","end-1c")
                self.case_num = int(self.case_num)
                del_and_open()


        
 
            Button(self.case_screen, text="Show the Story", width=16, height=1, command = lambda: retrieve_input() ).pack()

            


            
        root.after(2900, lambda : get_case_num())
        root.after(3000, root.destroy)

    def get_name(self):
      return self.name
  
    def get_age(self):
      return self.age

    def get_accompanied_by(self):
      return self.accompanied_by

    def get_location(self):
      return self.location

    def get_category_logic(self):
      self.category_logic = []
      tot_cat = [self.category_1,self.category_2,self.category_3,self.category_4,self.category_5,self.category_6,self.category_7,self.category_8,self.category_9]
      for cat in tot_cat:
        if cat!=cat:
          pass
        else:
          self.category_logic.append(cat)
      to_return = ""
      if len(self.category_logic) > 1:
        for inx,catu in enumerate(self.category_logic):
          if inx == len(self.category_logic) - 1:
            to_return = to_return[:-2] + " and " + catu
          else:
            to_return = to_return + " " + catu + ", "
      else:
        to_return = self.category_logic[0]
          


      return to_return


    def get_category_1(self):
      return self.category_1
    
    def get_category_2(self):
      return self.category_2

    def get_category_3(self):
      return self.category_3

    def get_category_4(self):
      return self.category_4

    def get_category_5(self):
      return self.category_5
    def get_category_6(self):
      return self.category_6

    def get_category_7(self):
      return self.category_7

    def get_category_8(self):
      return self.category_8

    def get_category_9(self):
      return self.category_9

    
    def seq_attempt(self):
      self.data_map = {}
      attempt_sheet = self.get_attempt_sheet()
      #print(attempt_sheet)
      print(self.total_df)

      for idx,row in self.total_df.iterrows():

        try:
        
          if attempt_sheet[idx]==1:
            if row['Sequence']!=row['Sequence']:
              pass
            else:
              self.data_map[str(row['Sequence'])] = row['Sub-feature']
        except:
          pass

      key_data_map = self.data_map
      key_data_map = sorted(key_data_map)
      #print(self.data_map)
      keys_list = key_data_map
      self.seq_to_index_map = {}
      self.index_to_seq_map = {}
      for inx,key in enumerate(keys_list):
        self.seq_to_index_map[key] = inx
        self.index_to_seq_map[inx] = key



    def check_for_extra_logic(self,query):

      if 'mul' in query[1:]: 
        if query[0] == '*':
          query_ls = query.split()
          
          if query_ls[0][1:] == "mul":
            for inx,let in enumerate(query):
              if let=='{':
                inx_s = inx
              elif let == '}':
                inx_e = inx
                break
              else:
                pass
          text = query[inx_s+1:inx_e]
      
          query_new = query.split('Range')
          query_new = query_new[-1]
          query_new = query_new[1:len(query_new)-1]
          #print(query_new)
          ranges_or = []
          ranges_and = []

          for inx,let in enumerate(query_new):
            if let == '[':
              range_start = inx
            elif let == ']':
              range_end = inx
              if ' & ' in query_new:
                t_query = query_new[range_start+1:range_end]
                ran_seq = t_query.split(':')
                range_1 = ran_seq[0].rstrip()
                range_1 = range_1.lstrip()
                range_2 = ran_seq[1].rstrip()
                range_2 = range_2.lstrip()
                ranges_and.append([range_1,range_2])
              elif ' or ' in query_new:
                t_query = query_new[range_start+1:range_end]
                ran_seq = t_query.split(':')
                range_1 = ran_seq[0].rstrip()
                range_1 = range_1.lstrip()
                range_2 = ran_seq[1].rstrip()
                range_2 = range_2.lstrip()
                ranges_or.append([range_1,range_2])
              else:
                t_query = query_new[range_start+1:range_end]
                ran_seq = t_query.split(':')
                range_1 = ran_seq[0].rstrip()
                range_1 = range_1.lstrip()
                range_2 = ran_seq[1].rstrip()
                range_2 = range_2.lstrip()
                range_let = [range_1,range_2]
            else:
              pass
            

          if len(ranges_and)>0 :
              sent = ""
      
              for ranges in ranges_and:
                
        
                range_st = int(self.seq_to_index_map[ranges[0]])
                range_en = int(self.seq_to_index_map[ranges[1]])
            
                for key in range(range_st,range_en + 1):
            
                  if self.data_map.get(self.index_to_seq_map[key])!= None:
                    
                    if key == range_en:
                      sent = sent + self.data_map[self.index_to_seq_map[key]] + ' and '
                    else:
                      sent =  sent + self.data_map[self.index_to_seq_map[key]] + ","
                  else:
                    pass
              

            

          elif len(ranges_or)>0 :
              sent = ""
      
              for ranges in ranges_and:
                
        
                range_st = int(self.seq_to_index_map[ranges[0]])
                range_en = int(self.seq_to_index_map[ranges[1]])
            
                for key in range(range_st,range_en + 1):
            
                  if self.data_map.get(self.index_to_seq_map[key])!= None:
                    
                    if key == range_en:
                      sent = sent + self.data_map[self.index_to_seq_map[key]] + ' and '
                    else:
                      sent =  sent + self.data_map[self.index_to_seq_map[key]] + ","
                  else:
                    pass

          else:
            sent = ""
          to_return = text + " " + sent
          to_return = to_return.rstrip()
          to_return = to_return.lstrip()
          return to_return
        
      elif 'make_plural' in query[1:]:
        p = inflect.engine()
        if query[0] == '*':
          query_ls = query.split()
        
          if query_ls[0][1:] == "make_plural":
            for inx,let in enumerate(query):
              if let=='{':
                inx_s = inx
              elif let == '}':
                inx_e = inx
                break
              else:
                pass
          text = query[inx_s+1:inx_e]
          query_new = query.split(':')
          query_new = query_new[-1]
          query_new = query_new[1:len(query_new)-1]
          params_norm = []
          params_or = []
          params_and = []

          if 'or' in query_new:
            query_list = query_new.split('or')
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_or.append(param_t)



          
          elif '&' in query_new:
            query_list = query_new.split('&')
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_and.append(param_t)


          else:

            query_list = query_new.split()
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_norm.append(param_t)
          check = 0
          if len(params_and)>0 :
            for param in params_and:
              if self.data_map.get(param)!=None:
                check = 1
                
              else:
                check = 0
                break

            if check == 1:
      
              text = p.plural(text)


          elif len(params_or)>0:
            for param in params_or:
              if self.data_map.get(param)!=None:
                check = 1
                break
              else:
                check = 0
              

            if check == 1:
          
              text = p.plural(text)
            
          elif len(params_norm)>0:
              if self.data_map.get(params_norm[0])!=None:
                check = 1
              else:
                check = 0
          

              if check == 1:
                text = p.plural(text)
            
          else:

            pass


        
          text = text.rstrip()
          text = text.lstrip()

          return text
      elif 'appearance' in query[1:]:
        p = inflect.engine()
        if query[0] == '*':
          query_ls = query.split()
        
          if query_ls[0][1:] == "appearance":
            for inx,let in enumerate(query):
              if let=='{':
                inx_s = inx
              elif let == '}':
                inx_e = inx
                break
              else:
                pass
          text = query[inx_s+1:inx_e]
          query_new = query.split(':')
          query_new = query_new[-1]
          query_new = query_new[1:len(query_new)-1]
          params_norm = []
          params_or = []
          params_and = []

          if 'or' in query_new:
            query_list = query_new.split('or')
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_or.append(param_t)



          
          elif '&' in query_new:
            query_list = query_new.split('&')
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_and.append(param_t)


          else:

            query_list = query_new.split()
            
            for queries in query_list:
              param_t = queries.rstrip()
              param_t = param_t.lstrip()
              params_norm.append(param_t)
          check = 0
          if len(params_and)>0 :
            for param in params_and:
              if self.data_map.get(param)!=None:
                check = 1
                
              else:
                check = 0
                break

            if check == 1:
      
              text=text
            else:
              text = ""
  


          elif len(params_or)>0:
            for param in params_or:
              if self.data_map.get(param)!=None:
                check = 1
                break
              else:
                check = 0
              

            if check == 1:
          
              text = text
            else:
              text = ""

  
          elif len(params_norm)>0:
              if self.data_map.get(params_norm[0])!=None:
                check = 1
              else:
                check = 0
          

              if check == 1:
                text = text
              else:
                text = ""

              
            
          else:

            text = ""


        

          text = text.rstrip()
          text = text.lstrip()
          return text


      else:
        return query

        

            






     

      
    def fetch_data(self):
      self.df_story_logic = pd.read_excel(self.story_logic_sheet)
      self.df_test_sheet = pd.read_excel(self.test_sheet)

      self.df_structure_logic = pd.read_excel(self.structure_logic_sheet)

      self.category_1 = self.df_structure_logic['Category-1'].values[0]
      self.category_2 = self.df_structure_logic['Category-2'].values[0]
      self.category_3 = self.df_structure_logic['Category-3'].values[0]
      self.category_4 = self.df_structure_logic['Category-4'].values[0]
      self.category_5 = self.df_structure_logic['Category-5'].values[0]
      self.category_6 = self.df_structure_logic['Category-6'].values[0]
      self.category_7 = self.df_structure_logic['Category-7'].values[0]
      self.category_8 = self.df_structure_logic['Category-8'].values[0]
      self.category_9 = self.df_structure_logic['Category-9'].values[0]


      self.clb_choice = self.df_structure_logic['Clubbing logic'].values

      self.rpt_choice = self.df_structure_logic['Report logic'].values

      self.addrs_choice = self.df_structure_logic['Address logic'].values
      #self.also_choice = df_structure_logic['Also logic']
      self.os_logic = self.df_structure_logic['OS logic'].values[0]
      self.location = self.df_structure_logic['Location'].values[0]
      self.also1_choice = self.df_structure_logic['Also-1 logic'].values
      self.also2_choice = self.df_structure_logic['Also-2 logic'].values

    def get_pronoun(self):
      if self.gender == "Female":
        return 'She'
      else:
        return 'He'
    
    def get_determiner(self):
      if self.gender == "Female":
        return 'her'
      else:
        return 'his'


    def get_os(self):
      """os_list = []
      for os_logic in self.os_logic:
        os_list.append(self.get_logic_sent(os_logic))
      print(os_list)
      """
      self.os = str(self.get_logic_sent(self.os_logic))

    def get_attempt_sheet(self):
      df_test_sheet = pd.read_excel(self.test_sheet)
      df2 = df_test_sheet[self.case_num]
      self.name = df2.iloc[1:2].values[0]
      self.age = df2.iloc[2:3].values[0]
      self.gender = df2.iloc[3:4].values[0]
      
      self.accompanied_by = df2.iloc[4:5].values[0]
      df2 = df2.fillna(0)
      self.df_test_sheet = df2.iloc[5:].values
      self.df_test_total = df2.iloc[5:].values
      

      self.row_lim = len(df_test_sheet)
      df_attempt = self.df_test_sheet
      print(df2.iloc[5:])
      df_attempt = np.where(df_attempt == 0,0,1)
      return df_attempt


    def make_me_plural(self,sent):
      sent = self.get_logic_sent(sent) + 's'
      return sent

    def get_dependency_logic(self,data_dict):

      if data_dict.get('B3-c') is not None or data_dict.get('B3-d') is not None:
        if data_dict.get('A-a') is not None:
          
          pass
        elif data_dict.get('A-b') is not None:
          
          pass

        else:
          pass


      else:
        pass

      return data_dict
        


    def get_total_df(self):
      df_story_logic = pd.read_excel(self.story_logic_sheet)

      self.total_df = df_story_logic.iloc[:self.row_lim,:]

    def get_category_sheet(self):
      category_sheet = self.total_df['Category'].values
      uniqueList = []
      for elem in category_sheet:
          if str(elem) not in uniqueList:
              uniqueList.append(str(elem))
      self.category_logic = ','.join(uniqueList)
      #category_sheet = list(map(int, category_sheet))
      return category_sheet

    def get_category_value(self):


      return str(self.category_sheet[self.idx])


    def get_data_logic_sheet(self):
      data_logic_sheet = self.df_test_total
      return data_logic_sheet


    def get_clubbing_sheet(self):
      df = self.total_df.fillna(0)
      df = df["Clubbing"].values
      clb_sheet = list(map(int, df))
      return clb_sheet

    def get_logic_sent(self,text):
      keys = self.struct_dict.keys()
      text = str(text)
      for key in keys:

        if key in text:
          if self.struct_dict.get(key)() is not None:
            text = text.replace('<'+key+'>',str(self.struct_dict.get(key)()))


      if '*' in text and '#' in text:
        
        in_st = text.index('*')
        in_en = text.index('#')
        text_before_that = text[:in_st]
        text_after_that = text[in_en+1:]

        text_to_pass = text[in_st : in_en]


        text_cr = self.check_for_extra_logic(text_to_pass)
        text_cr = text_cr.rstrip()
        text_cr = text_cr.lstrip()
        text = text_before_that + text_cr + " " + text_after_that
        text = text.rstrip()
        text = text.lstrip()

        
          



      return text


    def remove_nan(self,arr):
      new_arr = []
      for br in arr:
        if br!=br or br=="nan":
          pass
        else:
          new_arr.append(br)
      return new_arr


    def get_sequence_data(self):
      for idx,row in self.total_df.iterrows():
        self.struct_dict[str("data_") + str(row.Sequence)] = row['Sub-feature']


    def get_data_logic(self):

      return self.data_logic_sheet[self.idx]



    def get_report_logic(self):
      new_rpt_choice = []
      for ch in self.rpt_choice:
        if ch is not None:
          new_rpt_choice.append(self.get_logic_sent(ch))
      new_rept_choice = self.remove_nan(new_rpt_choice)

  
      return str(random.choice(new_rept_choice))


    def get_also1_logic(self):
      new_addrs_choice = []
      for ch in self.also1_choice:
        if ch is not None:
          new_addrs_choice.append(self.get_logic_sent(ch))


      new_ads_choice = self.remove_nan(new_addrs_choice)



      return str(random.choice(new_ads_choice))

    def get_also2_logic(self):
      new_addrs_choice = []
      for ch in self.also2_choice:
        if ch is not None:
          new_addrs_choice.append(self.get_logic_sent(ch))


      new_ads_choice = self.remove_nan(new_addrs_choice)



      return str(random.choice(new_ads_choice))

      



    def get_address_logic(self):
      new_addrs_choice = []
      for ch in self.addrs_choice:
        if ch is not None:
          new_addrs_choice.append(self.get_logic_sent(ch))


      new_ads_choice = self.remove_nan(new_addrs_choice)



      return str(random.choice(new_ads_choice))


    def get_clb_logic(self):
      new_clb_choice = []
      for ch in self.clb_choice:
        if ch is None:
          pass
        else:
          new_clb_choice.append(self.get_logic_sent(ch))


      new_clbt_choice = self.remove_nan(new_clb_choice)


    
      return str(random.choice(new_clbt_choice))

    def get_alphanumeric(self,key):
      return any(char.isdigit() for char in key)


    def pad_array(self,d,length):
      #print(len(np.pad(d, (0,(length - len(d)%length)), 'constant')))
      return np.pad(d, (0,(length - len(d)%length)), 'constant')

      

    def get_alpha(self):
      pass

    def get_dependency_sheet(self):
      def get_standard_matrix():
        filepath = self.story_logic_sheet

        sf = StyleFrame.read_excel(filepath , read_style=True, use_openpyxl_styles=False)




        def only_cells_with_red_text(cell):

            
            if cell.style.bg_color in {utils.colors.red, 'FFFF0000'}:
                return 120
        


        
        sf_2 = StyleFrame(sf.applymap(only_cells_with_red_text))





        #print(qualifying_dict)


        sf_2.to_excel().save()
        df = pd.read_excel(curdir+'/output.xlsx')


        df = df.iloc[:self.row_lim,1]
        #print(df)

        standard_matrix = df.values


        
        return standard_matrix
      
      sm = get_standard_matrix()
      sm = np.where(sm == 120,1,0)

      return sm

    def get_remark_sheet(self):
      def get_standard_matrix():
        filepath = self.story_logic_sheet

        sf = StyleFrame.read_excel(filepath , read_style=True, use_openpyxl_styles=False)




        def only_cells_with_red_text(cell):

            
            if cell.style.bg_color in {utils.colors.red, 'FFFF0000'}:
                return 120
        


        
        sf_2 = StyleFrame(sf.applymap(only_cells_with_red_text))





        #print(qualifying_dict)


        sf_2.to_excel().save()
        df = pd.read_excel(curdir+'/output.xlsx')


        df = df.iloc[:,3]
        #print(df)

        standard_matrix = df.values


        
        return standard_matrix
      
      sm = get_standard_matrix()
      sm = np.where(sm == 120,1,0)

      return sm
                              





    def get_struct_base(self,sre):
      sre = sre.replace("<","")
      sre = sre.split(">")
      structure = []
      for sr in sre:
        if sr == "":
          pass
        else:
          structure.append(sr)

      return structure



    def get_clubbing_logic(self,list_of_keys):
      
      club_pass = ""
      if len(list_of_keys) > 1: 
        for inx,key in enumerate(list_of_keys):
          if inx != len(list_of_keys) -1:
            club_pass = club_pass + key + ", "
          else:
            club_pass = club_pass[:-2] + " or " + key + "."
      else:
        club_pass = list_of_keys[0]


        
      clb_sent =  str(self.get_clb_logic()) + " " + str(club_pass)
      return clb_sent

    def final_processing(self,text):
      
      rx = r"\.(?=\S)"
      s = text
      result = re.sub(rx, ". ", s)
      return result
      
    def get_data_dict(self):
      clb_sent = []
      data_dict = {}

      total_sent = ""
      self.idx  = 1
      back_category_value = self.get_category_value()
      for self.idx,row in self.total_df.iterrows():
        if self.get_category_value()!=None:
          if back_category_value == self.get_category_value():
        
            
            if self.remark_sheet[self.idx] == 0:
              if self.df_attempt[self.idx] == 1:
                if self.clb_sheet[self.idx] == 1:
                  if row['AN-Data dictionary']!=row['AN-Data dictionary']:
                    pass
                  else:
                    clb_sent.append(self.get_logic_sent(str(row['AN-Data dictionary'])))
                  
                else:

                  if row.Sequence!=row.Sequence:
                    pass
                  else:

                    if self.get_alphanumeric(str(row.Sequence)):
      
                      
                        data_dict[row.Sequence] = self.get_logic_sent(str(row['AN-Data dictionary']))
                    
                  
                


                        

                        
                      #print(row.Sequence)
                      
                    else:
                      data_dict[row.Sequence] = "."+self.get_logic_sent(str(row['A-Sentence']))+"."
      
                    

              
              else:
                pass
            else:
              pass
          else:
            
            back_category_value = self.get_category_value()
            

            for key,value in data_dict.items():
              data_dict[key] = self.get_logic_sent(str(value))
            data_dict = self.get_dependency_logic(data_dict)
            

            keys_values = data_dict.items()

            data_dict = {str(key): str(value) for key, value in keys_values}
            




            
            







            total_sent = total_sent + self.get_my_sentence(data_dict,clb_sent) + "\n\n"



            clb_sent = []
            data_dict = {}
    



      for key,value in data_dict.items():
        data_dict[key] = self.get_logic_sent(str(value))
      data_dict = self.get_dependency_logic(data_dict)
      
      keys_values = data_dict.items()

      data_dict = {str(key): str(value) for key, value in keys_values}

      


      total_sent = total_sent + self.get_my_sentence(data_dict,clb_sent) 



            


      return total_sent



    def form_sentence(self,data_dict,clb_sent):
      with open("sentence_structure.txt") as tweetfile:
        sentstruct = json.load(tweetfile)
      sent = ""
      for i in range(1,5):
        if sentstruct['address_logic'] == i:
          sent = sent + "." + " " + self.get_address_logic() + " "
        elif sentstruct['report_logic'] == i:
          sent = sent + self.get_report_logic() + " "

        elif sentstruct['data_logic'] == i:
          for key,value in data_dict.items():
            sent = sent + str(value) + " "
        elif sentstruct['clubbing_logic'] == i:
          sent = sent + "." + clb_sent + "."
        else:
          pass




      return sent

    def gramarize(self,sent):

      f = open('Value-json/logic_activation.json') 
      activation = json.load(f)
      if activation['grammar_logic'] == "active":
        test_str = sent
        splitter = SentenceSplitter(language='en')
        sente = splitter.split(text=test_str)
        gram_sent = []

        for sent in sente:
          parser = GingerIt()

          output = parser.parse(sent)
          output_1 = (output.get("result"))
          output_1 = output_1 
          gram_sent.append(output_1)

        f_output = ' '.join(gram_sent)


        if f_output[-1] == '.' and f_output[-2] == '.':
          f_output = f_output[:-2]




        
        f_output = f_output + '.'

        f_output = self.remove_trailing_dots(f_output)
        f_output = f_output.replace('..','.')

        return f_output


      else:
        return sent
    




    def remove_trailing_dots(self,f_output):


      if f_output[-1] == '.' and f_output[-2] == '.':
        f_output = f_output[:-2]




      
      f_output = f_output + '.'

    

      return f_output



    def get_my_sentence(self,data_dict,clb_sent):
        
      data_dict = collections.OrderedDict(sorted(data_dict.items()))
    
      if len(clb_sent) > 0:
        clb_sent = self.get_clubbing_logic(clb_sent)
      else:
        clb_sent = ""

      sent = self.form_sentence(data_dict,clb_sent)

      sent = sent.replace("nan","")

      

    
      

      
      return str(self.gramarize(sent))

              
          
          

    

    def read_google_sheet(self):

        def read_path_info():

            file = open("google_path_info.txt","r")
            for lines in file.read().splitlines():
                if lines[0] == "T":
                    self.test_sheet = lines[1:]
                else:
                    self.story_logic_sheet = lines[1:]
            file.close() 
        
        def read_from_google_drive(filelink,destination):
            def get_id_from_link(filelink):
                stre = str(filelink)
                return stre[32:-17]

            def download_file_from_google_drive(id, destination):
                URL = "https://docs.google.com/uc?export=download"

                session = requests.Session()

                response = session.get(URL, params = { 'id' : id }, stream = True)
                token = get_confirm_token(response)

                if token:
                    params = { 'id' : id, 'confirm' : token }
                    response = session.get(URL, params = params, stream = True)

                save_response_content(response, destination)    

            def get_confirm_token(response):
                for key, value in response.cookies.items():
                    if key.startswith('download_warning'):
                        return value

                return None

            def save_response_content(response, destination):
                CHUNK_SIZE = 32768

                with open(destination, "wb") as f:
                    for chunk in response.iter_content(CHUNK_SIZE):
                        if chunk: # filter out keep-alive new chunks
                            f.write(chunk)


            file_id = get_id_from_link(filelink)
            download_file_from_google_drive(file_id, destination)

        read_path_info()
        
        read_from_google_drive(self.story_logic_sheet,'story_logic_sheet.xlsx')
        read_from_google_drive(self.test_sheet,'test_sheet.xlsx')
        self.display_user_options_window()



            
        

            


            


      

    def execute(self):
        
        self.display_user_options_window()
        #self.read_google_sheet()

        
    







            

    

    


    
        


   
def scanfiles():
    global scan_screen
    scan_screen = Tk()
    curdir = os.getcwd()

    bg_path = curdir+"\my_bg.png"
    #C = Canvas(main_screen, bg="blue", height=1024, width=1024)
    img = ImageTk.PhotoImage(Image.open(bg_path))
    ico_path = curdir+"\myicon.ico"
    scan_screen.iconbitmap(ico_path)
    scan_screen.geometry("1024x1021")
    scan_screen.title("Scan files from Google sheet")
    Label(image=img, width="1024", height="300").pack()
    Label(text="").pack()
    def retrieve_input():

    
        delete_login_success()
        

    photo_login = PhotoImage(file = curdir+"\scandoc.png") 
    
    
    # Resizing image to fit on button 
    #photo_login = photo_login.subsample(5, 5)
    #photo_register = photo_register.subsample(5, 5)  
    
    # here, image option is used to 
    # set image on button 
    # compound option is used to align 
    # image on LEFT side of button 
    def get_story_logic_sheet():
        story_logic_sheet = tkinter.filedialog.askopenfilename(parent=scan_screen, initialdir=curdir,title='Please Choose Story Logic Sheet',filetypes=[('Excel File', '.xlsx'),('CSV Excel file', '.csv')])
        story_logic_sheet = story_logic_sheet.rstrip()
        file1 = open("google_path_info.txt", "w")
        file1.write("L"+story_logic_sheet)
        file1.write(" \n")
        file1.close()

    def get_test_sheet():
        test_sheet = tkinter.filedialog.askopenfilename(parent=scan_screen, initialdir=curdir,title='Please Choose Test Logic Sheet',filetypes=[('Excel File', '.xlsx'),('CSV Excel file', '.csv')])
        test_sheet = test_sheet.rstrip()
        file1 = open("google_path_info.txt", "a")
        file1.write("T"+test_sheet)
        file1.write(" \n")
        file1.close()     
    """
    Label(text="Paste Google Logic Sheet Link").pack()
    textBox=Text(scan_screen, height=2, width=40)
    
    textBox.pack()
    Label(text="Paste Google Test Input Sheet Link").pack()
    textBox1=Text(scan_screen, height=2, width=40)
    
    textBox1.pack()
    Label(text="").pack()
    """
    photo_login = PhotoImage(file = curdir+"\login.png") 
    photo_scan = PhotoImage(file = curdir+"\scandoc.png") 
    photo_register = PhotoImage(file = curdir+"\\register.png") 
    
    # Resizing image to fit on button 
    #photo_login = photo_login.subsample(5, 5)
    #photo_register = photo_register.subsample(5, 5)  
    
    # here, image option is used to 
    # set image on button 
    # compound option is used to align 
    # image on LEFT side of button 
    Button(text = '    Select Story Logic Sheet', height="60", width="200",  image = photo_login, 
                        compound = LEFT, command = get_story_logic_sheet).pack() 
    
    Label(text="").pack()
    Button(text = '     Select Test Sheet', height="60", width="200",  image = photo_register, 
                        compound = LEFT, command = get_test_sheet).pack() 
    Label(text="").pack()
    Button(text = '    Scan it!!', height="60", width="180",  image = photo_scan, 
                        compound = LEFT, command = lambda: retrieve_input()).pack() 
    
    Label(text="").pack()
    

    #C.pack()
 
    scan_screen.mainloop()

scanfiles()

