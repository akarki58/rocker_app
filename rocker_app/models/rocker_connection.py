# -*- coding: utf-8 -*-
#############################################################################
#
#    Copyright (C) 2019-Antti Kärki.
#    Author: Antti Kärki.
#
#    You can modify it under the terms of the GNU AFFERO
#    GENERAL PUBLIC LICENSE (AGPL v3), Version 3.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU AFFERO GENERAL PUBLIC LICENSE (AGPL v3) for more details.
#
#    You should have received a copy of the GNU AFFERO GENERAL PUBLIC LICENSE
#    (AGPL v3) along with this program.
#    If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################


from odoo import api, fields, models
from odoo import exceptions
import logging

_logger = logging.getLogger(__name__)


class rocker_connection():

    # @api.multi
    #
    def create_connection(self):

        _database_record = self
        _datasource = _database_record.name
        _driver = _database_record.driver
        _sqlalchemydriver = _database_record.sqlalchemydriver
        _odbcdriver = _database_record.odbcdriver
        _sid = _database_record.database
        _database = _database_record.database
        _host = _database_record.host
        _port = _database_record.port
        _user = _database_record.user
        _password = _database_record.password

        con = None
        engine = None
        _logger.info('Connecting to database: ' + _database)
        try:
            if _driver == 'postgresql':
                _logger.info('Using PostgreSQL')
                import psycopg2
            elif _driver == "sqlalchemy":
                _logger.info('Using SQLAlchemy')
                import sqlalchemy
            elif _driver == "mysql":
                _logger.info('Using MySQL')
                import mysql.connector
            elif _driver == "mariadb":
                _logger.info('Using MariaDB')
                import mysql.connector
            elif _driver == "oracle":
                _logger.info('Using Oracle')
                # _logger.debug(_user + '/' + _password + '@' + _host + ':' + _port + '/' + _sid)
                import cx_Oracle
            elif _driver == "sqlserver":
                _logger.info('Using SQLServer')
                import pyodbc
            elif _driver == "odbc":
                _logger.info('Using ODBC')
                import pyodbc
            else:
                raise exceptions.ValidationError('Driver not supported')
        except ModuleNotFoundError as moduleErr:
            print("[Error]: Failed to import (Module Not Found) {}.".format(moduleErr.args[0]))
            raise exceptions.ValidationError(
                "[Error]: Failed to import (Module Not Found) {}.".format(moduleErr.args[0]))
            sys.exit(1)
        except ImportError as impErr:
            print("[Error]: Failed to import (Import Error) {}.".format(impErr.args[0]))
            raise exceptions.ValidationError("[Error]: Failed to import (Import Error) {}.".format(impErr.args[0]))
            sys.exit(1)

        try:
            if _driver == 'postgresql':
                try:
                    con = psycopg2.connect(host=_host, port=_port, database=_database, user=_user, password=_password)
                    cursor = con.cursor()
                except (Exception, psycopg2.Error) as error:
                    print(f"Error connecting to the database: {error}")
                    sys.exit(1)
                finally:
                    if cursor:
                        cursor.close()
                        # connection.close()
                        # print("Database connection closed.")
            elif _driver == "sqlalchemy":
                try:
                    engine = sqlalchemy.create_engine(f"{_sqlalchemydriver}://{_user}:{_password}@{_host}:{_port}/{_database}")
                    con = engine.raw_connection()
                except Exception as err:
                    raise exceptions.ValidationError('SQLAlchemy drivers\n'+err)
            elif _driver == "mysql":
                con = mysql.connector.connect(host=_host, port=_port, database=_database, user=_user,
                                              password=_password)
            elif _driver == "mariadb":
                con = mysql.connector.connect(host=_host, port=_port, database=_database, user=_user,
                                              password=_password)
            elif _driver == "oracle":
                _logger.debug('Try Oracle')
                # _logger.debug(_user + '/' + _password + '@' + _host + ':' + _port + '/' + _sid)
                try:
                    # cx_Oracle.init_oracle_client()
                    con = cx_Oracle.connect(_user + '/' + _password + '@' + _host + ':' + _port + '/' + _sid)
                except Exception as err:
                    _logger.debug("Whoops in Oracle connect!")
                    _logger.debug(err)
                    raise exceptions.ValidationError('Oracle drivers\n'+err)
            elif _driver == "sqlserver":
                # _logger.debug(
                #     'DRIVER={' + _odbcdriver + '};SERVER=' + _host + ';DATABASE=' + _database + ';UID=' + _user + ';PWD=' + _password)
                con = pyodbc.connect(
                    'DRIVER={' + _odbcdriver + '};SERVER=' + _host + ';DATABASE=' + _database + ';UID=' + _user + ';PWD=' + _password)
                self._sqldriver = 'sqlserver'
            elif _driver == "odbc":
                # _logger.debug(
                #     'DRIVER={' + _odbcdriver + '};SERVER=' + _host + ';DATABASE=' + _database + ';UID=' + _user + ';PWD=' + _password)
                con = pyodbc.connect(
                    'DRIVER={' + _odbcdriver + '};SERVER=' + _host + ';DATABASE=' + _database + ';UID=' + _user + ';PWD=' + _password)
                self._sqldriver = 'odbc'
            else:
                raise exceptions.ValidationError('Driver not supported')
        except:
            _logger.debug('Database connection failed')
            raise exceptions.ValidationError('Database connection failed')

        _logger.debug('Connection: ')
        _logger.debug(con)
        _logger.debug('Engine: ')
        _logger.debug(engine)
        if _driver == "sqlalchemy":
            return engine, con
        else:
            return con
