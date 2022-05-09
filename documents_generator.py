import pandas as pd
import datetime
import easygui

# getting file locations
start_directory = easygui.enterbox(r"Please provide source file path. Example: C:\Users\User1234\Desktop\data.csv")

final_directory = easygui.enterbox(r'Please provde final file path with extension ".html". Example: C:\Users\User1234\Desktop\final_document.html')

# assigning date
today = f'Szczebrzeszyn, day of {datetime.datetime.today().strftime("%d-%m-%Y")}'

# reading data
df = pd.read_csv(start_directory,skiprows=0,encoding='utf-8',sep=';')

# adjusting the format of a column with salaries
df['Wyn. zasadn.'] = df['Wyn. zasadn.'].str.replace(' ','').str.replace(',','.').astype(float)

# adding column with name and last name and removing separate ones
df['osoba'] = df['Imię'] + " " + df['Nazwisko']
df = df.drop(columns=['Nazwisko','Imię'])
# setting first line of a document
first_line = f'''
<div class="page">
 <div class="first_line">
            <section class="court"> <p>SĄD PONADOKRĘGOWY</p> <p>W SZCZEBRZESZYNIE</p> <p>Koziołka Matołka 5</p> <p>99-656 Szczebrzeszyn</p></section>
            <section class="place_n_time"> {today} r.</section>
        </div>
'''

# replacing NA cells with empty string
df.fillna('',inplace=True)

# setting up dictionaries for indicating salaries in words (i.e. 1500 should be followed by "One thousand five hundred zlotych")
thousand = {1:'tysiąc',2:'dwa tysiące',3:'trzy tysiące',4:"cztery tysiące",5:"pięc tysięcy",6:"sześć tysięcy",7:'siedem tysięcy',8:"osiem tysięcy",9:'dziewięć tysięcy'}
hundred = {0:'',1:'sto',2:'dwieście',3:'trzysta',4:'czterysta',5:'pięćset',6:"sześćset",7:'siedemset',8:'osiemset',9:'dziewięćset'}
ten = {0:'',2:'dwadzieścia',3:'trzydzieści',4:'czterdzieści',5:'pięćdziesiąt',6:'sześćdziesiąt',7:'siedemdziesiąt',8:'osiemdziesiąt',9:'dziewięćdziesiąt'}
exceptions = {0:'dziesięć',1:'jedenaście',2:"dwanaście",3:'trzynaście',4:"czternaście",5:"piętnaście",6:'szesnaście',7:'siedemnaście',8:'osiemnaście',9:'dziewiętnaście'}
single = {0:'',1:'jeden',2:'dwa',3:'trzy',4:'cztery',5:'pięć',6:'sześć',7:'siedem',8:'osiem',9:'dziewięć'}

# setting up function to get salary as string of words
def print_text_salary(x):
    x = str(x)
    print(x)
    if int(x[-4]) == 1:
        return thousand[int(x[-6])]+' ' + hundred[int(x[-5])] + ' ' + exceptions[int(x[-3])] +' zł'   
    elif int(x[-4]) == 0:
        return thousand[int(x[-6])] + " " + hundred[int(x[-5])]  + " " + single[int(x[-3])] + ' zł'
    else:
        return thousand[int(x[-6])] + " " + hundred[int(x[-5])] + " " + ten[int(x[-4])] + " " + single[int(x[-3])] + ' zł'

# beggining of the official text with settings for page in css
start_text=f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
    <style>
 body {{
    margin: 0;
    padding: 0;
    background-color: #FAFAFA;
    font: 12pt "Tahoma";
}}
* {{
    box-sizing: border-box;
    -moz-box-sizing: border-box;
}}
.page {{
    width: 21cm;
    min-height: 29.7cm;
    padding: 2cm;
    margin: 1cm auto;
    border: 1px #D3D3D3 solid;
    border-radius: 5px;
    background: white;
    box-shadow: 0 0 5px rgba(0, 0, 0, 0.1);
}}
.subpage {{
    padding: 1cm;
    border: 2px red solid;
    height: 237mm;
    outline: 2cm #FFEAEA solid;
}}

.first_line{{
display: flex;
flex-direction: row;
justify-content: space-between;
}}

.second_line{{
    display: flex;
    flex-direction: row;
    justify-content: flex-end;
    margin-bottom: 120px;
    margin-top: 80px;
}}
section.court {{
    width:5cm;

}}
section.court p {{ 
    margin:0;
    text-align: center;
 }}

 section.signature p {{
     margin:0;
 }}

 section.person_title{{
     width:8cm;
 }}


 section.person_title p{{
     margin:0;
     text-align: left;
 }}

 section.official_text{{
      text-align: justify;
 }}

 section.signature{{
     margin-top: 10cm;
 }}
@page {{
    size: A4;
    margin: 0;
}}
@media print {{
    .page {{
        margin: 0;
        border: initial;
        border-radius: initial;
        width: initial;
        min-height: initial;
        box-shadow: initial;
        background: initial;
        page-break-after: always;
    }}
}}
</style>
</head>
<body>

'''

# end of text

end_text = '''
</body>
</html>
'''

# starting the process
full_text = ''
full_text += start_text

# looping over each person in a list and generating a text
for x in df.itertuples():
    # check if woman or man
    if x[3] == "K":
        title = 'Pani'
        title_inside = 'Pani'
    else:
        title = 'Pan'
        title_inside = 'Panu'

    # check if funkcja exists
    if len(x[4]) == 0:
        second_line = f'''
        <div class="second_line">
            <section class="person_title"> <p>{title} {x[6]}</p>
            <p>{x[2]}</p>
            <p>Sądu Okręgowego w Łodzi</p> </section>
        </div>
        '''
    else:
        second_line = f'''
        <div class="second_line">
            <section class="person_title"> <p>{title} {x[6]}</p>
            <p>{x[2]}</p>
            <p>{x[4]}</p>
            <p>Sądu Ponadkręgowego w Szczebrzeszynie</p> </section>
        </div>
        '''
    official_text = f'''
    <section class="official_text"> &emsp;Na podstawie rozporządzenia Ministra Sprawiedliwości z dnia 3 marca 2017 r. w sprawie stanowisk i 
        szczegółowych zasad wynagradzania urzędników i innych pracowników sądów&nbsp; i prokuratury 
        oraz odbywania stażu urzędniczego (Dz. U. z 2017 r., poz. 485 z późn. zm.)
        <p>&emsp; -przyznaję {title_inside} z <b>dniem 1 czerwca 2022 r. wynagrodzenie zasadnicze</b> w kwocie <b>{"{:.2f}".format(x[5])} zł</b> ({print_text_salary(x[5]).replace('  ',' ')}) miesięcznie.</p>
        <p> &emsp;Pozostałe składniki wynagrodzenia pozostają bez zmian.</p>
        </section>
        <section class="signature"> <u>Do wiadomości i wykonania:</u>
        <p>Kier. Oddz. Finansowego SP w Szczerzbrzeszynie</p>
            </section>
    </div>
    </div>
    '''
    full_text = full_text + first_line + second_line + official_text

full_text +=end_text
  

# saving the file in html
with open(final_directory, "a",encoding="UTF-8") as f:
    f.write(full_text)
