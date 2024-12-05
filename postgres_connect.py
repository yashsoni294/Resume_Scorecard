import psycopg2

hostname = 'localhost'
username = 'postgres'
password = '123456'
database = 'ResumeDB'
port_id = 5432

conn = None
cur = None

try:
    conn = psycopg2.connect(
        host=hostname,
        user=username,
        password=password,
        dbname=database,
        port=port_id
    )

    cur = conn.cursor()

    create_script = """ CREATE TABLE IF NOT EXISTS resumes_table (
        resume_id SERIAL PRIMARY KEY,
        resume_file_name VARCHAR(255),
        resume_file_text TEXT,
        resume_key_aspect TEXT,
        resume_score FLOAT
    ); """

    cur.execute(create_script)
    conn.commit()




    cur.close()
    conn.close()

except (Exception, psycopg2.Error) as error:
    print("Error while connecting to PostgreSQL", error)

finally:
    if cur is not None:
        cur.close()
        print('Cursor closed.')

    if conn is not None:
        conn.close()
        print('Database connection closed.') 