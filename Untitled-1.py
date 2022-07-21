from lib2to3.pgen2.token import OP
from typer import Option

from Talisman import Talisman
from SkillList import SkillList
from Misc import Misce
from tkinter import *
from tkinter import ttk
import pandas as pd
import openpyxl
from openpyxl import *


path='Talisman Chart.xlsx'
root = Tk()
root.geometry("600x100")
root.title("Talisman Entry")

df = pd.read_excel(path)
df = df.reset_index(drop=True)

skills = SkillList.skills
skills.sort()
rarList = Misce.rarityRange
levelRange = Misce.skillLevelRange
decoLevels = Misce.decoRange

def genTalisman():
    global df
    gen = [rarClick.get(),s1Click.get(),s1LevClick.get(),s2Click.get(),
    s2LevClick.get(),dec1Click.get(),dec2Click.get(),dec3Click.get()]
    df2 = pd.DataFrame([gen], columns=["Rarity", "Skill 1", "Skill 1 Level", 
    "Skill 2", "Skill 2 Level", "Slot 1", "Slot 2", "Slot 3"])
    df = pd.concat([df,df2],ignore_index=True)
    print(df)
    with pd.ExcelWriter(path) as writer:
        df.to_excel(writer, sheet_name='new_sheet2',index = False)
  
Label(root,text="Rarity").grid(row=1,column=1)
rarClick = StringVar()
rarClick.set("-")
rarityDrop = OptionMenu(root, rarClick, *rarList)
rarityDrop.grid(row=2,column=1)

Label(root,text="Skill 1").grid(row=1,column=2)
s1Click = StringVar(root)
s1Click.set("--------")
skillDrop1 = ttk.Combobox(root, textvariable = s1Click, values = skills)
skillDrop1.grid(row=2, column=2)

Label(root,text="Skill 1 Lvl").grid(row=1,column=3)
s1LevClick = StringVar()
s1LevClick.set("-")
s1LevDrop = OptionMenu(root, s1LevClick, *levelRange)
s1LevDrop.grid(row=2,column=3)

Label(root,text="Skill 2").grid(row=1,column=4)
s2Click = StringVar(root)
s2Click.set("--------")
skillDrop2 = ttk.Combobox(root, textvariable = s2Click, values = skills)
skillDrop2.grid(row=2, column=4)

Label(root,text="Skill 2 Lvl").grid(row=1,column=5)
s2LevClick = StringVar()
s2LevClick.set("-")
s2LevDrop = OptionMenu(root, s2LevClick, *levelRange)
s2LevDrop.grid(row=2,column=5)

Label(root,text="Deco 1").grid(row=1,column=7)
dec1Click = StringVar()
dec1Click.set("-")
dec1Drop = OptionMenu(root, dec1Click, *decoLevels)
dec1Drop.grid(row=2,column=7)

Label(root,text="Deco 2").grid(row=1,column=8)
dec2Click = StringVar()
dec2Click.set("-")
dec2Drop = OptionMenu(root, dec2Click, *decoLevels)
dec2Drop.grid(row=2,column=8)

Label(root,text="Deco 3").grid(row=1,column=9)
dec3Click = StringVar()
dec3Click.set("-")
dec3Drop = OptionMenu(root, dec3Click, *decoLevels)
dec3Drop.grid(row=2,column=9)

Label(root,text="Ready to enter?").grid(row=3,column=4)
finishedButton = Button(root, text = "Enter",command = genTalisman).grid(row=4,column=4)
root.mainloop()
