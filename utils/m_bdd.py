#! /bin/python3

import psycopg2

# Permet de se connecter à une base postgres
def db_connect(host_name, user_name, user_password, db_name):
    connection = None
    try:
        connection = psycopg2.connect(
            user = user_name,
            password = user_password,
            host = host_name,
            database = db_name
        )
        print("BDD connected")
    except:
        print("Erreur de Connexion à la base de données :", db_name)
        pass
    return connection

# Permet de lancer une requête sur une base postgres
def db_query(connection, query, type = 'select'):
    cursor = connection.cursor()
    result = None
    try:
        cursor.execute(query)
        if type == 'select':
            result = cursor.fetchall()
            return result
    except:
        print("Erreur de requete.")
        pass

