import pdf2docx
from kivymd.app import MDApp
from kivy.lang import Builder
from plyer import filechooser
kv = """
FloatLayout:
    MDLabel:
        text:"PDF-WORD CONVERTER"
        halign:"center"
        font_size:32
        font_name:"Poppins-Light"
        pos_hint : {'center_x':0.5,'center_y':0.8}
        size_hint_y:None
        height:100
        size_hint_x:None
        width:300

    MDRaisedButton:
        text:"CHOOSE FILE"
        font_name:"Poppins-Light"
        pos_hint : {'center_x':0.5,'center_y':0.6}
        on_release:
            app.choose()
    MDLabel:
        text:"PLEASE SELECT A PDF FILE AND MAKE SURE THAT THE FILE NAME DOES NOT HAVE '.' IN IT"
        id:addedfile
        halign:"center"
        font_name:"Poppins-Light"
    MDRaisedButton:
        text:"CONVERT FILE"
        font_name:"Poppins-Light"
        pos_hint : {'center_x':0.5,'center_y':0.4}
        on_release:
            app.convert()
    MDLabel:
        id:process
        text:""
        halign:"center"
        pos_hint : {'center_x':0.5,'center_y':0.3}
        font_name:"Poppins-Light"

"""
class converter(MDApp):
    def build(self):
        return Builder.load_string(kv)
    def choose(self):
        filechooser.open_file(on_selection = self.selected)
    def selected(self,selection):
        self.root.ids.process.text = ""
        global pdf
        pdf = selection
        if pdf != []:
            
            path = pdf[0].split("\\")
            if path[len(path)-1].split(".")[1] != "pdf":
                self.root.ids.addedfile.text = "PLEASE UPLOAD PDF FILES ONLY"

            else:
                self.root.ids.addedfile.text = pdf[0] 
        else:
            pass
    def convert(self):
        global pdf
        if(pdf):
            self.root.ids.process.text = "PROCESSING"
            path = pdf[0].split("\\")
            wordfilename = path[len(path)-1].split(".")[0]
            word = f"{wordfilename}-converted.docx"
            cv = pdf2docx.Converter(pdf[0])
            cv.convert(word,start=0,end = None)
            cv.close()
            self.root.ids.process.text = "FILE SUCCESSFULLY CONVERTED!!\nFILE WILL BE SAVED IN THE SAME DIRECTORY AS THAT OF SOURCE FILE"
        else:
            self.root.ids.process.text = "PLEASE SELECT A FILE"
        
pdf = ""
converter().run()