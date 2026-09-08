import os
import pandas as pd
import asana
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from sqlalchemy import MetaData
import win32com.client

try:
    # Loads variables from a local .env file (if present) into os.environ.
    # This file is git-ignored, so real secrets never get committed.
    # If python-dotenv isn't installed, we just rely on real environment
    # variables being set some other way (e.g. system/user env vars).
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


def scriptmonitoring(
    script="scriptname",
    scriptstart="1/1/1899",
    scriptfinish="1/1/1899",
    scriptinput="",
    scriptoutput="T",
    user="ipoole",
    _uuid="",
    result="",
    failure=0,
):
    """This scriptmonitoring requires a script,
    the start time, end time, script input,
    script output, the user (which should be defined via the
    ScriptImportStart), a result, and whether or not the script succeeded
    or failed reprsented as 1 or a 0
    """
    # setting up asana for later usage
    asana_access_token = os.environ.get("CIDA_ASANA_ACCESS_TOKEN")
    if not asana_access_token:
        raise EnvironmentError(
            "CIDA_ASANA_ACCESS_TOKEN environment variable is not set. "
            "Set it before running this script (see .env.example)."
        )
    configuration = asana.Configuration()
    configuration.access_token = asana_access_token
    api_client = asana.ApiClient(configuration)
    tasks_api_instance = asana.TasksApi(api_client)
    opts = {}

    ScriptMonitoringUpdate = pd.DataFrame(
        columns=[
            "script",
            "starttime",
            "endtime",
            "scriptinput",
            "scriptoutput",
            "user",
            "uuid",
            "result",
        ],
        data=[
            [
                script,
                scriptstart,
                scriptfinish,
                scriptinput,
                scriptoutput,
                user,
                _uuid,
                result,
            ]
        ],
    )
    # finally we run the data either to the SQL server or if ti fails to get to the sql server as a backup to a csv
    try:
        # this relates to the production issues log
        project = "1203415330244683"
        server = "CIDA-SQL-T-01"
        database = "cida_montioring"
        username = os.environ.get("CIDA_SQL_USERNAME")
        password = os.environ.get("CIDA_SQL_PASSWORD")
        if not username or not password:
            raise EnvironmentError(
                "CIDA_SQL_USERNAME and/or CIDA_SQL_PASSWORD environment "
                "variables are not set. Set them before running this "
                "script (see .env.example)."
            )
        table = "python_script_monitoring"
        connection_string = (
            "DRIVER={ODBC Driver 13 for SQL Server};SERVER="
            + server
            + ";DATABASE="
            + database
            + ";UID="
            + username
            + ";PWD="
            + password
        )
        connection_url = URL.create(
            "mssql+pyodbc", query={"odbc_connect": connection_string}
        )
        # despit the fact we should only be adding one we still use fast_executemany cause the developer is lazy
        engine = create_engine(connection_url, fast_executemany=True)
        cnxn = engine.connect()
        meta = MetaData()
        meta.reflect(engine)
        datatable = meta.tables[table]
        # again, we want to make sure our types are set based on what's in the sql server, so this grabs those type sand then fixes the datetimes
        # when we aren't able to do so we'll make datetimes later in the queries
        pytypedict = {}
        sqltypedict = {}
        pycolumns = []
        for c in datatable.columns:
            pytypedict[c.name] = c.type.python_type
            sqltypedict[c.name] = c.type
            pycolumns += [c.name]
        ScriptMonitoringUpdate.columns = pycolumns
        pytypedict.update({"starttime": "datetime64[ns]", "endtime": "datetime64[ns]"})
        ScriptMonitoringUpdate = ScriptMonitoringUpdate.astype(pytypedict)
        ScriptMonitoringUpdate.to_sql(
            table,
            engine,
            if_exists="append",
            index=False,
            chunksize=100,
            dtype=sqltypedict,
        )
        cnxn.close()
    #    print('sqlsuccess')
    # the below is so that if we fail to write to SQL we isntead right to a backup csv
    except:
        try:
            cnxn.close()
        except:
            pass
        os.chdir(
            r"C:\Users"
            + "\\"
            + user
            + r"\Boston University\Continuous Improvement & Data Analytics - Documents\Automation and Data\MonitoringForCIDADesktop"
        )
        ScriptMonitoringCSV = pd.read_csv("ScriptMonitoringCSV.csv")
        ScriptMonitoringCSV = pd.concat([ScriptMonitoringCSV, ScriptMonitoringUpdate])
        ScriptMonitoringCSV.to_csv("ScriptMonitoringCSV.csv", index=False)

        try:
            # this relates to the bu.edu workspace and may need to be changed in the future, but we can also create the projet with just the ID
            # workspace = '676371944581295'
            # the below creates the task in the project based on project id and assigns it to ipoole, best practice would be to assign it to whomever owns the script/automation
            # client.tasks.create_task({'name':'Test','assignee':'ipoole@bu.edu','projects':project})
            body = {
                "data": {
                    "name": script + " failed in SQL Monitoring",
                    "notes": scriptoutput,
                    "assignee": "ipoole@bu.edu",
                    "projects": project,
                }
            }
            tasks_api_instance.create_task(body, opts)
        except:
            emailsubject = "CIDASCRIPTLOG: " + script + " failed"
            emailbody = script + " failed in SQL Monitoring & Asana"
            outlook = win32com.client.Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = "cidadata@bu.edu"
            mail.Subject = emailsubject
            mail.Body = emailbody
            mail.Send()
    if failure == 1:
        # the very simple asana script below builds out a client and then creates the task
        try:
            body = {
                "data": {
                    "name": script + " failed",
                    "notes": scriptoutput,
                    "assignee": "ipoole@bu.edu",
                    "projects": project,
                }
            }
            tasks_api_instance.create_task(body, opts)
        # if that breaks we send an email to cidadata@bu.edu
        except:
            emailsubject = "CIDASCRIPTLOG: " + script + " failed"
            emailbody = (
                user
                + ","
                + script
                + ","
                + scriptstart
                + ","
                + scriptfinish
                + ","
                + scriptinput
                + ","
                + scriptoutput
                + ", failed to reach Asana"
            )
            outlook = win32com.client.Dispatch("outlook.application")
            mail = outlook.CreateItem(0)
            mail.To = "cidadata@bu.edu"
            mail.Subject = emailsubject
            mail.Body = emailbody
            mail.Send()
