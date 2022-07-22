''' init.py '''

# Libraries
import mysql.connector

# Credenciais database
credentials = {"user": "","password": "","host": "","port": "","database": ""}

def conn_db():
    execute_query_response = {}
    execute_query_response['status'] = False
    execute_query_response['content'] = None
    conn = None
    try:
        conn = mysql.connector.connect(user=credentials['user'],password=credentials['password'],host=credentials['host'],port=credentials['port'],database=credentials['database'])
    except Exception as error:
        execute_query_response['content'] = str(error)
    if conn is not None:
        execute_query_response['status'] = True
        execute_query_response['content'] = conn
    return execute_query_response

def execute_query_connection(cur=None,content=''):
    execute_query_response= {}
    execute_query_response['status']= False
    execute_query_response['content']= None
    try:
        if 'SELECT' in content:
            cur.execute(content)
            fetchall_response = cur.fetchall()
            execute_query_response['content'] = fetchall_response
            execute_query_response['status'] = True
        elif 'INSERT' in content or 'UPDATE' in content:
            cur.execute(content)
            execute_query_response['content'] = 'Sucess in Insert/Update'
            execute_query_response['status'] = True
    except Exception as error:
        sql = ''
        try:
            sql = content
        except Exception:
            sql = str(content) + ' : '
            pass
        execute_query_response['content'] = str(error) + ' : ' + str(sql)
        execute_query_response['status']= False
    return execute_query_response