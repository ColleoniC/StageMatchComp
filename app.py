from flask import Flask, render_template, redirect, url_for
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import declarative_base, sessionmaker
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

obiettivo = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credenziali = ServiceAccountCredentials.from_json_keyfile_name('Credenziali.json', obiettivo)
client = gspread.authorize(credenziali)

engine = create_engine("sqlite:///database.db")
Base = declarative_base()

class Questionario(Base):
    __tablename__ = 'questionari'
    id = Column(Integer, primary_key=True, autoincrement=True)
    Indirizzo_email = Column(String(255), unique=True, nullable=False)
    Nome = Column(String(100), nullable=False)
    Cognome = Column(String(100), nullable=False)
    Luogo_Di_Residenza = Column(String(100))
    Interesse_1 = Column(Integer)
    Interesse_2 = Column(Integer)
    Interesse_3 = Column(Integer)
    Interesse_4 = Column(Integer)
    Interesse_5 = Column(Integer)
    Interesse_6 = Column(Integer)
    Interesse_7 = Column(Integer)
    Interesse_8 = Column(Integer)
    Interesse_9 = Column(Integer)
    Interesse_10 = Column(Integer)
    Interesse_11 = Column(Integer)
    Interesse_12 = Column(Integer)
    Interesse_13 = Column(Integer)
    Interesse_14 = Column(Integer)
    Interesse_15 = Column(Integer)
    Interesse_16 = Column(Integer)
    Interesse_17 = Column(Integer)
    Interesse_18 = Column(Integer)
    Interesse_19 = Column(Integer)
    Competenza_1 = Column(Integer)
    Competenza_2 = Column(Integer)
    Competenza_3 = Column(Integer)
    Competenza_4 = Column(Integer)
    Competenza_5 = Column(Integer)
    Competenza_6 = Column(Integer)
    Competenza_7 = Column(Integer)
    Competenza_8 = Column(Integer)
    Competenza_9 = Column(Integer)
    Competenza_10 = Column(Integer)
    Competenza_11 = Column(Integer)
    Competenza_12 = Column(Integer)
    Competenza_13 = Column(Integer)
    Competenza_14 = Column(Integer)
    Competenza_15 = Column(Integer)
    Competenza_16 = Column(Integer)
    Competenza_17 = Column(Integer)
    Competenza_18 = Column(Integer)
    Competenza_19 = Column(Integer)
    
    def __getitem__(self, key):
        return getattr(self, key)

Base.metadata.create_all(engine)
Session = sessionmaker(bind=engine)

def sync_google_sheets():
    session = Session()
    try:
        sheet = client.open("Questionario PCTO - Studente").sheet1
        data = sheet.get_all_records()
        
        for riga in data:
            email = riga.get('Indirizzo email', '').strip()
            if not email:
                continue
            
            esistente = session.query(Questionario).filter_by(Indirizzo_email=email).first()
            
            dati = {
                'Indirizzo_email': email,
                'Nome': riga.get('Nome', '').strip(),
                'Cognome': riga.get('Cognome', '').strip(),
                'Luogo_Di_Residenza': riga.get('Luogo di residenza', '').strip()
            }
            
            for i in range(1, 20):
                val_int = riga.get(f'Interesse {i}')
                val_com = riga.get(f'Competenza {i}')
                dati[f'Interesse_{i}'] = int(val_int) if val_int not in ('', None) else None
                dati[f'Competenza_{i}'] = int(val_com) if val_com not in ('', None) else None
            
            if esistente:
                for key, value in dati.items():
                    setattr(esistente, key, value)
            else:
                session.add(Questionario(**dati))
            
            session.commit()
            
    except Exception as e:
        session.rollback()
        print(f"Errore sync: {e}")
    finally:
        session.close()

@app.route('/')
def home():
    sync_google_sheets()
    session = Session()
    try:
        questionario = session.query(Questionario).first()
        if questionario:
            return redirect(url_for('visualizza', id=questionario.id))
        return render_template('error.html', messaggio='Nessun questionario trovato')
    finally:
        session.close()

@app.route('/visualizza/<int:id>')
def visualizza(id):
    sync_google_sheets()
    session = Session()
    try:
        questionario = session.query(Questionario).filter_by(id=id).first()
        if not questionario:
            return render_template('error.html', messaggio='Questionario non trovato')
        return render_template('index.html', questionario=questionario)
    finally:
        session.close()

if __name__ == '__main__':
    print("Server su http://127.0.0.1:5000")
    app.run(host='127.0.0.1', port=5000, debug=True)