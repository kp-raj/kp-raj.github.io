import openpyxl, random
from openpyxl import Workbook
from flask import Flask, render_template, request

app = Flask(__name__)

def leaderboard():
    names=[]
    scores=[]
    for i in range(2,12):
        names.append(sheet_board['B'+str(i)].value)
        scores.append(sheet_board['C'+str(i)].value)
    return names,scores


wb=openpyxl.load_workbook('lyrics.xlsx')
sheet_board=wb['Board']
score=0
top_name, top_score=leaderboard()
lang=''
ans=''
col=''
mode_g=''

@app.route('/', methods=['GET','POST'])
def index():
    return render_template('home.html', top_score=top_score, top_name=top_name, score=score)

@app.route('/mode', methods=['GET','POST'])
def mode():
    global lang
    language=request.args.get('language')
    if language=='':
        language=lang
    lang=language

    return render_template('mode.html',language=language,top_score=top_score, top_name=top_name, score=score)

@app.route('/game',methods=['GET','POST'])
def game():
    mode=request.args.get('GameMode')
    global lang,col, ans, mode_g
    sheet=wb[lang]
    if mode=='':
        mode=mode_g
    if mode=='Song':
        col='C'
    elif mode=='Lyricist':
        col='D'
    elif mode=='Movie':
        col='E'
    else:
        col=random.choice(['C','D','E'])[0]
    mode_g=mode
    count=sheet.max_row
    lyr=random.randrange(2,count+1)
    lyric=(sheet['G'+str(lyr)].value).split(';')
    game_guess=sheet[col+'1'].value
    ans=sheet[col+str(lyr)].value
    options=[ans]
    while len(options)<4:
        option=sheet[col+str(random.randrange(2,count+1))].value
        if option not in options:
            options.append(option)
    options.sort()

    
    return render_template('game.html',lyric=lyric,game_guess=game_guess,options=options,top_score=top_score, top_name=top_name, score=score)

@app.route('/check')
def check():
    answer=request.args.get('answer')
    global ans, score, col, mode_g
    add=0
    answer=' '.join(answer.split("+"))
    if answer==ans:
        flag=True
        if col=='D':
            add=2
        else:
            add=1
        score=score+add
    else:
        flag=False
    return render_template('result.html',flag=flag,top_score=top_score, top_name=top_name, score=score, answer=ans,add=add,mode=mode_g)

@app.route('/exit')
def exiting():
    save=request.args.get('save')
    name=request.args.get('name')
    if name=='':
        name='Shy'
    global score, sheet_board, top_name, top_score
    bigger=2
    if save=='Yes':
        if score>sheet_board['C11'].value:
            for i in range(2,12):
                if score>sheet_board['C'+str(i)].value:
                    sheet_board.insert_rows(i)
                    sheet_board['B'+str(i)]=name
                    sheet_board['C'+str(i)]=score
                    break
        wb.save('lyrics.xlsx')
        top_name,top_score=leaderboard()
    score=0
    return render_template('home.html', top_score=top_score, top_name=top_name, score=score)

@app.errorhandler(404)
def page_not_found(e):
    return render_template('error.html',top_score=top_score, top_name=top_name, score=score)
     

        


if __name__=='__main__':
    app.run()