# 03.02.2025 - 24.05.2025 

try:
    import importlib
    import subprocess
    import sys
    import os
    import pandas as pd
    import sqlite3 
    import datetime
    import time
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.backends.backend_pdf import PdfPages
    import seaborn as sns
    import math

except ImportError as e:
    print(f'Error during import: {e}')


class gym():

    '''Class variables needed for the communication between the methods.'''
    directory = None
    connection = None 
    cursor = None
    fileset = set()
    set_database = set()
    data = None
    df = None


    @classmethod
    def install(cls):

        '''Automatically installs all necessary modules.'''
        modules = ['numpy', 'pandas', 'matplotlib']
        for module in modules:
            try:
                importlib.import_module(module) 
                # print(f'{module} is alredy installed.')
            except ImportError:
                print(f'{module} is not installed. Installing...')
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', module])
        print('Intallations checked.')


    @classmethod
    def set_directory(cls):
        
        try:
            cls.directory = os.path.dirname(os.path.abspath(__file__))
            # print(cls.directory)
            # directory = '..' # navigates to the previous directory
            # directory = '.' # calls the current directory
        except Exception as e:
            print(f'Error during set_directory: {e}')


    @classmethod
    def connect(cls):
        
        '''Ensures a proper conection to the Database.'''
        try:
            with sqlite3.connect('Gym-Database.db') as connection:
                cls.connection = connection    
                cls.cursor = cls.connection.cursor()
                # print('Database connected.')
                return cls.connection, cls.cursor
        except Exception as e:
            print(f'Error during connect: {e}')

    
    @classmethod
    def close_con(cls):
        
        '''Manually close the connection to the database.'''
        try:
            if cls.connection is not None:
                cls.connection.commit() # 'commit' hat in SQL eine besondere 
                cls.connection.close()  # Bedeutung. Stichwort Roleback!
                cls.connection = None
                cls.cursor = None
                print('Database closed.')     
        except Exception as e:
            print(f'Error during close_con: {e}')

    
    @classmethod
    def wait(cls):

        try:
            time.sleep(2) 
        except Exception as e:
            print(f'Error during wait: {e}')


    @classmethod
    def create_table(cls):
        
        '''Creates the table of the database'''
        try:
            if not hasattr(cls, 'connection') or cls.connection is None: # Prüfen! Wurde noch nie ausgelöst...
                raise ValueError('Connection not established')  
            cls.cursor.execute('''
            CREATE TABLE IF NOT EXISTS Gym_Plan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            Datum TEXT,
            Körpergewicht INTEGER,
            Reihenfolge INTEGER,
            Wiederholungen TEXT,
            Übung TEXT,
            Gewicht INTEGER,
            Satz1 INTEGER,
            Satz2 INTEGER,
            Satz3 INTEGER,
            Satz4 INTEGER,
            Bemerkung TEXT
            )
            ''')                                                  
            cls.connection.commit()
            print('Table created (or already exists).')
        except Exception as e:
            print(f'Error during create_table: {e}')


    @classmethod
    def create_fileset(cls):
        
        '''Creates a list of the files in the directory of 
        the Python script to turn them into a DataFrame.'''
        content = os.listdir(cls.directory)
        try:
            for file in content:
                # print(file)
                if file not in cls.fileset: 
                    if file.endswith('.xlsx') and not file.startswith('~'):
                        cls.fileset.add(file)                    
                    else:
                        continue
                else: 
                    continue 
            print(f'Fileset created: {cls.fileset}') 
        except Exception as e:
            print(f'Error during create_fileset: {e}')

    
    @classmethod
    def create_set_database(cls):
        '''Creates a set with all the dates of the database.'''

        query_date = 'SELECT Datum FROM Gym_Plan'
        try:
            cls.set_database.clear()
            query = cls.cursor.execute(query_date)
            dates = query.fetchall()
            column_names = [description[0] for description in query.description]
            df = pd.DataFrame(dates, columns=column_names)
            for date in df['Datum']:
                if date is not None and not pd.isna(date):
                    formatted_date = cls.format_date(date)
                    formatted_filename = f'{formatted_date}.xlsx'
                    if formatted_filename not in cls.set_database:
                        cls.set_database.add(formatted_filename)
                else:
                    continue
            print(f'set_database created: {cls.set_database}')
        except Exception as e:
            print(f'Error during dates: {e}')


    @classmethod
    def format_date(cls, date_value):
        
        '''Ensures a uniform date format.'''
        if isinstance(date_value, datetime.datetime):
            return date_value.strftime('%Y.%m.%d')
        return str(date_value).replace('-', '.')
    
        
    @classmethod
    def fill_database(cls):

        '''Inserts data from Excel tables into database.'''

        try: 
            for file in cls.fileset:
                date_part = file.replace('.xlsx', '')
                formatted_date = cls.format_date(date_part)
                formatted_filename = f'{formatted_date}.xlsx'
                if file not in cls.set_database: 
                    file_path = os.path.join(cls.directory, file)
                    df = pd.read_excel(file_path)
                    print(f'Inserting {file} into database...')
                    df['Datum'] = df['Datum'].apply(cls.format_date)
                    for row in df.itertuples(index=False): # index False vs. True
                        try:
                            cls.cursor.execute('''
                            INSERT INTO Gym_Plan (
                            Datum,
                            Körpergewicht,
                            Reihenfolge, 
                            Wiederholungen, 
                            Übung, 
                            Gewicht,
                            Satz1, 
                            Satz2, 
                            Satz3, 
                            Satz4, 
                            Bemerkung
                            )
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', row)
                        except Exception as e:
                            print(f'Error while inserting row: {e}')
                    cls.set_database.add(formatted_filename)
                    cls.connection.commit()
                    print(f'File {formatted_filename} inserted.')
                    print(f'Check set_database: {cls.set_database}')
                else: 
                    print(f'File {file} already in database. Skipping file.')
                    continue
        except Exception as e:
            print(f'Error during fill_database: {e}')

    
    @classmethod
    def reset_database(cls):
        '''Resets the database and fills it up again.'''
        try:
            confirmation = input('''
            Are you sure you want to reset the database? (y/n): ''')
            if confirmation.lower() == 'y':
                cls.cursor.execute('DELETE FROM Gym_Plan')
                cls.connection.commit()
                cls.set_database.clear()
                cls.fill_database()
                print('Database has been reset.')
            else:
                print('Process canceled.')

        except Exception as e:
            print(f'Error while reseting the database: {e}')


    @classmethod
    def clean_database(cls):

        '''Checks the directory for missing Excel files and deletes the corresponding data 
        records from the database to dajust the database to the status of the directory .'''
        
        print('Checking for outdated data...')
        try:
            for dataset in cls.set_database:
                if dataset in cls.set_database and dataset not in cls.fileset:
                    file_name = dataset.split('.')[:-1]
                    formatted_date = '-'.join(file_name)
                    try:
                        cls.cursor.execute('DELETE FROM Gym_Plan WHERE Datum = ?', (formatted_date,))
                        print(f'Deleted {dataset} from database.')
                    except Exception as e:
                        print(f'Error while cleaning database: File: {file}, Error: {e}')
                else:
                    continue
            cls.connection.commit()
            print('Database is clean.')
        except Exception as e:
            print(f'Error during clean_database: {e} ')


    @classmethod
    def read_database(cls):

        '''Reads the entiere database and returns a 
        dataframe which will be used for data-analyses.'''
        query_all = 'SELECT * FROM Gym_Plan'
        try:
            query = cls.cursor.execute(query_all)
            data = query.fetchall()
            column_names = [description[0] for description in query.description]
            df = pd.DataFrame(data, columns=column_names)
            cls.df = df
        except Exception as e:
            print(f'Error read_database: {e}')


    @classmethod
    def data_analyses(cls, filename='Progress.pdf'):
        '''Creates a graph of the progression for each exercise and saves it as a PDF.'''
        try:
            # 1. Daten aus Datenbank auslesen. Check.
            query = 'SELECT Datum, Übung, Gewicht FROM Gym_Plan'
            cls.cursor.execute(query)
            data = cls.cursor.fetchall()
            df = pd.DataFrame(data, columns=['Datum', 'Übung', 'Gewicht'])
            df['Datum'] = pd.to_datetime(df['Datum'])
    
            # Daten vorbereiten
            df.set_index('Datum', inplace=True)
            drop_ex = ['Laufen (Gewicht = km/h, Satz = min)',
                       'Schultern aufgewärmt? (0 = nein / 1 = ja)',
                       'Rotatorenmanschette Rechts, innen',
                       'Rotatorenmanschette Links, innen',
                       'Rotatorenmanschette Rechts, außen',
                       'Rotatorenmanschette Links, außen']
    
            df_drop = df[~df['Übung'].isin(drop_ex)]
            df_ready = df_drop.groupby(['Datum', 'Übung'])['Gewicht'].max().reset_index()
            df_ready.set_index('Datum', inplace=True)
            
            exercises = df_ready['Übung'].unique()
            num_exercises = len(exercises)
            start_date = df_ready.index.min()
            end_date = df_ready.index.max()
            num_cols = 2
            num_rows = math.ceil(num_exercises / num_cols)
        
            fig, axes = plt.subplots(num_rows, num_cols, figsize=(11.69, 8.27), sharex=True)  # A4-Größe in Zoll
            fig.suptitle('Progress', fontsize=16)
    
            date_format = mdates.DateFormatter('%Y-%m-%d')
            plt.gca().xaxis.set_major_formatter(date_format)
            plt.gca().xaxis.set_major_locator(mdates.AutoDateLocator())
            plt.gcf().autofmt_xdate()
            
            for i, exercise in enumerate(exercises):
                row = i // num_cols
                col = i % num_cols
                ax = axes[row, col] if num_rows > 1 else axes[col]
        
                df_exercise = df_ready[df_ready['Übung'] == exercise]
                sns.lineplot(data=df_exercise, x=df_exercise.index, y='Gewicht', ax=ax)
                ax.set_title(exercise)
                ax.set_ylabel('Gewicht (kg)')
                ax.grid(True)
        
            for i in range(num_exercises, num_rows * num_cols):
                row = i // num_cols
                col = i % num_cols
                fig.delaxes(axes[row, col] if num_rows > 1 else axes[col])
        
            plt.tight_layout()
            plt.show()
        except Exception as e:
            print(f'Error during data_analyses {e}')
        
        try:
            with PdfPages(filename) as pdf:
                pdf.savefig(fig)
            print(f'Diagram successfully saved as a PDF file: {filename}')
        except Exception as e:
            print(f'Error during PDF creation: {e}')
        finally:
            plt.close(fig)



if __name__ == '__main__':

    gym.install()
    gym.set_directory()
    gym.connect()
    gym.wait()
    gym.create_table()   
    gym.create_fileset()
    gym.create_set_database()
    gym.fill_database()   
    gym.clean_database()
    gym.read_database()

    # Activate only if you want to reset the database!
    # Please disable again after using!
    #gym.reset_database()
    
    gym.data_analyses()

    gym.close_con()


