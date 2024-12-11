import psycopg2
from datetime import datetime

hostname = 'localhost'
username = 'postgres'
password = '123456'
database = 'ResumeDB'
port_id = 5432

conn = None
cur = None

def convert_To_Binary(filename): 
    with open(filename, 'rb') as file: 
        data = file.read() 
    return data 

try:
    conn = psycopg2.connect(
        host=hostname,
        user=username,
        password=password,
        dbname=database,
        port=port_id
    )

    cur = conn.cursor()

    cur.execute("""
            CREATE TABLE IF NOT EXISTS resume_table (
                unique_id NUMERIC PRIMARY KEY,
                resume_name VARCHAR(100) ,
                resume_content TEXT ,
                resume_key_aspect TEXT ,
                score INTEGER ,
                blob_data BYTEA 
            )
        """)
    conn.commit()
    unique_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
    file_data = convert_To_Binary(r"extracted_files\20241209102637721686_Naukri_AmitSinghal[13y_0m](1).pdf") 
    resume_content = "This is a test resume"
    resume_name = "Naukri_AmitSinghal[13y_0m](1).pdf"
    # resume_key_aspect = "This is a test resume"
    # score = 100
    # BLOB DataType 
    BLOB = psycopg2.Binary(file_data) 
  
    # SQL query to insert data into the database. 
    cur.execute( 
            "INSERT INTO resume_table(unique_id,resume_name,resume_content,blob_data) "
            "VALUES(%s,%s,%s,%s)", (unique_id, resume_name, resume_content, BLOB)
            )

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

def retrieve_resume_blob(unique_id):
    conn = None
    try:
        # Connect to the PostgreSQL server
        conn = psycopg2.connect(
            host=hostname,
            dbname=database,
            user=username,
            password=password,
            port=port_id
        )

        # Create a cursor
        cur = conn.cursor()

        # SQL query to retrieve the BLOB data
        cur.execute(
            "SELECT unique_id, resume_name, blob_data FROM resume_table WHERE unique_id = %s", 
            (unique_id,)
        )

        # Fetch the result
        result = cur.fetchone()

        if result:
            unique_id, resume_name, blob_data = result
            
            # Define the output directory (you can modify this path as needed)
            output_dir = r"extracted_files"
            
            # Create the directory if it doesn't exist
            import os
            os.makedirs(output_dir, exist_ok=True)

            # Full path for the output file
            output_path = os.path.join(output_dir, f"{unique_id}_{resume_name}")

            # Write the BLOB data to a file
            with open(output_path, 'wb') as file:
                file.write(blob_data)

            print(f"Resume retrieved and saved to: {output_path}")
            return output_path
        else:
            print(f"No resume found with unique_id: {unique_id}")
            return None

    except (Exception, psycopg2.DatabaseError) as error:
        print(f"Error retrieving resume: {error}")
    finally:
        if cur is not None:
            cur.close()
            print('Cursor closed.')

        if conn is not None:
            conn.close()
            print('Database connection closed.')

    return output_path

# Example usage
print(retrieve_resume_blob("20241210161750106808"))


def update_resume_details(unique_id, resume_key_aspect, score):
    conn = None
    try:
        # Connect to the PostgreSQL server
        conn = psycopg2.connect(
            host=hostname,
            dbname=database,
            user=username,
            password=password,
            port=port_id
        )

        # Create a cursor
        cur = conn.cursor()

        # SQL query to update resume_key_aspect and score
        cur.execute(
            """
            UPDATE resume_table 
            SET resume_key_aspect = %s, 
                score = %s 
            WHERE unique_id = %s
            """, 
            (resume_key_aspect, score, unique_id)
        )

        # Commit the changes
        conn.commit()

        print(f"Updated details for unique_id: {unique_id}")

    except (Exception, psycopg2.DatabaseError) as error:
        print(f"Error updating resume details: {error}")
    finally:
        if cur is not None:
            cur.close()
            print('Cursor closed.')

        if conn is not None:
            conn.close()
            print('Database connection closed.')

# Example usage
# update_resume_details(
#     unique_id="20241209150344045164", 
#     resume_key_aspect="Software Engineer with extensive Python and Machine Learning experience", 
#     score=85
# )