import datetime
import getpass
import glob
import ntpath
import os
import re
import sys
import time
import traceback
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import uuid
from datetime import date
from datetime import datetime as datet
from datetime import timedelta as td
from tkinter import *
from tkinter import filedialog


import lxml
import numpy as np
import xlrd
from sqlalchemy import text

import pandas as pd

failure = 0
user = getpass.getuser()
_uuid = str(uuid.uuid4())
scriptstart = datetime.datetime.now()
scriptstart = scriptstart.strftime("%Y-%m-%d %H:%M:%S")
def get_last_runtime(server, database, username, password, script, table, schema):
    # Creates connection to sql server and the meta engine that will allow us to access the tables
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
    connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": connection_string})
    engine = create_engine(connection_url, fast_executemany=True, future=True)
    cnxn = engine.connect()

    lastrun = pd.read_sql(
        "SELECT MAX(starttime) AS lastrun FROM "
        + schema
        + "."
        + table
        + " WHERE script = '"
        + script
        + "' AND (result = 'Success' OR result IS NULL)",
        engine,
    )
    lastruntime = lastrun["lastrun"][0]
    cnxn.close()
    
    return lastruntime



def joinstringsofdatframe(dataframe, row: int = 0):
    stringthing = ""
    for i in range(len(dataframe.columns)):
        try:
            if i != len(dataframe.columns) - 1:
                stringthing += str(dataframe.iloc[row][i]) + "|"
            else:
                stringthing += str(dataframe.iloc[row][i])
        except:
            pass
    return stringthing


def get_information(
    directory,
    onlyfolders=[],
    include_stats=False,
    include_colnames=False,
    only_files=False,
    file_end = '*',
):
    """This gets infomration from a directory, there a few things to note.
    
    Onlyfolders allows you to submit a list that will get just those folders read, especially importnat when using include_colnames as that can be resource intesive
    
    include_stats includes the numbe of files in the folder and the size of the folder
    
    include_colnames gets the first files' column names and stores them as a text string
    
    only_files flips this to just read files while the others are more for reading whole directories.
    
    file_end allows you to specify which file ending you want (like xls, doc, pdf, etc.) it defaults to just getting 'em all via *
    This is relevant just ofr only files

    """
    file_list = []
    if onlyfolders == []:
        foldermapping = os.listdir(directory)
    else:
        foldermapping = onlyfolders
    if only_files == True:
        glob_files = glob.glob(directory + '\\' + file_end)

        for current_file in glob_files:
            replaced = str.replace(current_file,directory + '\\','')
            filestats = os.stat(directory + '\\' + replaced)
            file_list.append(
                [
                    replaced,
                    time.ctime(filestats.st_mtime),
                    time.ctime(filestats.st_ctime),
                ]
            )
    else:
        for currentfolder in foldermapping:
            file_size = 0
            folderpath = os.path.join(directory, currentfolder)
            columns_in_datasource = (
                "Unable to match file to read, only text, excel and csv matched"
            )

            folderstats = os.stat(folderpath)
            if os.path.isfile(folderpath) == False:
                folderlist = os.listdir(folderpath)

                if include_stats == True or include_colnames == True:
                    for ele in os.scandir(folderpath):
                        file_size += os.path.getsize(ele)
                    folderlist = [
                        file
                        for file in folderlist
                        if file.endswith(".xls")
                        or file.endswith(".xlsx")
                        or file.endswith(".csv")
                        or file.endswith(".txt")
                        or file.endswith(".xlsm")
                    ]

                    folderlist = [
                        file
                        for file in folderlist
                        if os.path.isfile(os.path.join(folderpath, file)) == True
                    ]

                    file_num = len(os.listdir(folderpath))
                if include_colnames == True:
                    try:
                        if folderlist == []:
                            columns_in_datasource = "No files"
                        else:
                            # ariba files get a bit odd, with the utf-8 on the csvs and the
                            if currentfolder.startswith("Ariba") == True:
                                if (
                                    re.match(
                                        ".*\.xls$|.*\.xlsx$|.*\.xlsm$", folderlist[0]
                                    )
                                    != None
                                ):
                                    firstfile = pd.read_html(
                                        os.path.join(folderpath, folderlist[0]),
                                        header=None,
                                    )
                                    firstfile = firstfile[1]
                                    firstfile = pd.DataFrame(
                                        np.vstack([firstfile.columns, firstfile])
                                    )
                                    firstfile = firstfile.loc[[0]]
                                    columns_in_datasource = joinstringsofdatframe(
                                        firstfile
                                    )
                                if re.match(".*\.csv$", folderlist[0]) != None:
                                    try:
                                        firstfile = pd.read_csv(
                                            os.path.join(folderpath, folderlist[0]),
                                            nrows=10,
                                            header=None,
                                        )
                                    except Exception as e:
                                        newskip = 0
                                        try:
                                            if re.search("line\s\d", str(e)) != None:
                                                findalllist = re.findall("\d,", str(e))
                                                newskip = (
                                                    int(findalllist[0].replace(",", ""))
                                                    - 1
                                                )
                                                firstfile = pd.read_csv(
                                                    os.path.join(
                                                        folderpath, folderlist[0]
                                                    ),
                                                    skiprows=newskip,
                                                    nrows=10,
                                                    header=None,
                                                )
                                        except Exception as e:
                                            if re.search("line\s\d", str(e)) != None:
                                                try:
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                except Exception as e:
                                                    findalllist = re.findall(
                                                        "\d,", str(e)
                                                    )
                                                    newskip = (
                                                        int(
                                                            findalllist[0].replace(
                                                                ",", ""
                                                            )
                                                        )
                                                        - 1
                                                    )
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        skiprows=newskip,
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                    if (
                                                        pd.isnull(firstfile.iloc[0, 0])
                                                        == True
                                                    ):
                                                        firstfile = firstfile.iloc[
                                                            :, 1:
                                                        ]
                                                        firstfile = firstfile.iloc[[0]]
                                                    else:
                                                        first_valid_loc = firstfile.apply(
                                                            lambda col: col.first_valid_index()
                                                        ).max()
                                                        firstfile = firstfile.loc[
                                                            [first_valid_loc]
                                                        ]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                                    first_valid_loc = firstfile.apply(
                                        lambda col: col.first_valid_index()
                                    ).max()
                                    firstfile = firstfile.loc[[first_valid_loc]]
                                    columns_in_datasource = joinstringsofdatframe(
                                        firstfile
                                    )
                            elif (
                                currentfolder.startswith("SAP") == True
                                or currentfolder.startswith("QM4") == True
                            ):
                                if (
                                    re.match(
                                        ".*\.xls$|.*\.xlsx$|.*\.xlsm$", folderlist[0]
                                    )
                                    != None
                                ):
                                    firstfile = pd.read_excel(
                                        os.path.join(folderpath, folderlist[0]),
                                        nrows=10,
                                        header=None,
                                    )
                                    firstfile = firstfile.iloc[:, 1:]
                                    first_valid_loc = firstfile.apply(
                                        lambda col: col.first_valid_index()
                                    ).max()
                                    firstfile = firstfile.loc[[first_valid_loc]]
                                    columns_in_datasource = joinstringsofdatframe(
                                        firstfile
                                    )
                                if re.match(".*\.csv$", folderlist[0]) != None:
                                    try:
                                        firstfile = pd.read_csv(
                                            os.path.join(folderpath, folderlist[0]),
                                            nrows=10,
                                            header=None,
                                        )
                                        firstfile = firstfile.iloc[:, 1:]
                                        first_valid_loc = firstfile.apply(
                                            lambda col: col.first_valid_index()
                                        ).max()
                                        firstfile = firstfile.loc[[first_valid_loc]]
                                        columns_in_datasource = joinstringsofdatframe(
                                            firstfile
                                        )
                                    except Exception as e:
                                        newskip = 0
                                        try:
                                            if re.search("line\s\d", str(e)) != None:
                                                findalllist = re.findall("\d,", str(e))
                                                newskip = int(
                                                    findalllist[0].replace(",", "")
                                                )
                                                firstfile = pd.read_csv(
                                                    os.path.join(
                                                        folderpath, folderlist[0]
                                                    ),
                                                    skiprows=newskip,
                                                    nrows=10,
                                                    header=None,
                                                )
                                                firstfile = firstfile.iloc[:, 1:]
                                                first_valid_loc = firstfile.apply(
                                                    lambda col: col.first_valid_index()
                                                ).max()
                                                firstfile = firstfile.loc[
                                                    [first_valid_loc]
                                                ]
                                                columns_in_datasource = (
                                                    joinstringsofdatframe(firstfile)
                                                )
                                        except Exception as e:
                                            if re.search("line\s\d", str(e)) != None:
                                                try:
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                    firstfile = firstfile.iloc[:, 1:]
                                                    firstfile = firstfile.loc[[0]]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                                                except Exception as e:
                                                    findalllist = re.findall(
                                                        "\d,", str(e)
                                                    )
                                                    newskip = int(
                                                        findalllist[0].replace(",", "")
                                                    )
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        skiprows=newskip,
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                    firstfile = firstfile.iloc[:, 1:]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                                if re.match(".*\.txt$", folderlist[0]) != None:
                                    try:
                                        firstfile = pd.read_csv(
                                            os.path.join(folderpath, folderlist[0]),
                                            sep="\t",
                                            nrows=5,
                                            header=None,
                                        )
                                    except Exception as e:
                                        newskip = 0
                                        if re.search("line\s\d", str(e)) != None:
                                            findalllist = re.findall("\d,", str(e))
                                            newskip = (
                                                int(findalllist[0].replace(",", "")) - 1
                                            )
                                            firstfile = pd.read_csv(
                                                os.path.join(folderpath, folderlist[0]),
                                                sep="\t",
                                                skiprows=newskip,
                                                nrows=1,
                                                header=None,
                                            )
                                    firstfile = firstfile.loc[[0]]
                                    columns_in_datasource = joinstringsofdatframe(
                                        firstfile
                                    )
                            else:
                                if (
                                    re.match(
                                        ".*\.xls$|.*\.xlsx$|.*\.xlsm$", folderlist[0]
                                    )
                                    != None
                                ):
                                    firstfile = pd.read_excel(
                                        os.path.join(folderpath, folderlist[0]),
                                        nrows=10,
                                        header=None,
                                    )
                                    first_valid_loc = firstfile.apply(
                                        lambda col: col.first_valid_index()
                                    ).max()
                                    firstfile = firstfile.loc[[first_valid_loc]]
                                    columns_in_datasource = joinstringsofdatframe(
                                        firstfile
                                    )
                                if re.match(".*\.csv$", folderlist[0]) != None:
                                    try:
                                        firstfile = pd.read_csv(
                                            os.path.join(folderpath, folderlist[0]),
                                            nrows=10,
                                            header=None,
                                        )
                                        first_valid_loc = firstfile.apply(
                                            lambda col: col.first_valid_index()
                                        ).max()
                                        firstfile = firstfile.loc[[first_valid_loc]]
                                        columns_in_datasource = joinstringsofdatframe(
                                            firstfile
                                        )
                                    except Exception as e:
                                        newskip = 0
                                        try:
                                            if re.search("line\s\d", str(e)) != None:
                                                findalllist = re.findall("\d,", str(e))
                                                newskip = int(
                                                    findalllist[0].replace(",", "")
                                                )
                                                firstfile = pd.read_csv(
                                                    os.path.join(
                                                        folderpath, folderlist[0]
                                                    ),
                                                    skiprows=newskip,
                                                    nrows=10,
                                                    header=None,
                                                )
                                                first_valid_loc = firstfile.apply(
                                                    lambda col: col.first_valid_index()
                                                ).max()
                                                firstfile = firstfile.loc[
                                                    [first_valid_loc]
                                                ]
                                                columns_in_datasource = (
                                                    joinstringsofdatframe(firstfile)
                                                )
                                        except Exception as e:
                                            if re.search("line\s\d", str(e)) != None:
                                                try:
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                    firstfile = firstfile.loc[[0]]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                                                except Exception as e:
                                                    findalllist = re.findall(
                                                        "\d,", str(e)
                                                    )
                                                    newskip = int(
                                                        findalllist[0].replace(",", "")
                                                    )
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep=";",
                                                        skiprows=newskip,
                                                        nrows=10,
                                                        header=None,
                                                    )
                                                    firstfile = firstfile.loc[[0]]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                                if re.match(".*\.txt$", folderlist[0]) != None:
                                    try:
                                        firstfile = pd.read_csv(
                                            os.path.join(folderpath, folderlist[0]),
                                            sep="\t",
                                            nrows=5,
                                            header=None,
                                        )
                                        firstfile = firstfile.loc[[0]]
                                        columns_in_datasource = joinstringsofdatframe(
                                            firstfile
                                        )
                                    except Exception as e:
                                        try:
                                            newskip = 0
                                            if re.search("line\s\d", str(e)) != None:
                                                findalllist = re.findall("\d,", str(e))
                                                newskip = (
                                                    int(findalllist[0].replace(",", ""))
                                                    - 1
                                                )
                                                firstfile = pd.read_csv(
                                                    os.path.join(
                                                        folderpath, folderlist[0]
                                                    ),
                                                    sep="\t",
                                                    skiprows=newskip,
                                                    nrows=1,
                                                    header=None,
                                                )
                                                firstfile = firstfile.loc[[0]]
                                                columns_in_datasource = (
                                                    joinstringsofdatframe(firstfile)
                                                )
                                        except Exception as e:
                                            try:
                                                firstfile = pd.read_csv(
                                                    os.path.join(
                                                        folderpath, folderlist[0]
                                                    ),
                                                    sep=",",
                                                    nrows=5,
                                                    header=None,
                                                )
                                                firstfile = firstfile.loc[[0]]
                                                columns_in_datasource = (
                                                    joinstringsofdatframe(firstfile)
                                                )
                                            except Exception as e:
                                                newskip = 0
                                                if (
                                                    re.search("line\s\d", str(e))
                                                    != None
                                                ):
                                                    findalllist = re.findall(
                                                        "\d,", str(e)
                                                    )
                                                    newskip = (
                                                        int(
                                                            findalllist[0].replace(
                                                                ",", ""
                                                            )
                                                        )
                                                        - 1
                                                    )
                                                    firstfile = pd.read_csv(
                                                        os.path.join(
                                                            folderpath, folderlist[0]
                                                        ),
                                                        sep="\t",
                                                        skiprows=newskip,
                                                        nrows=1,
                                                        header=None,
                                                    )
                                                    firstfile = firstfile.loc[[0]]
                                                    columns_in_datasource = (
                                                        joinstringsofdatframe(firstfile)
                                                    )
                    except Exception as e:
                        columns_in_datasource = "Failed to get file columns"
                        pass
                    file_list.append(
                        [
                            currentfolder,
                            time.ctime(folderstats.st_mtime),
                            time.ctime(folderstats.st_ctime),
                            file_size,
                            file_num,
                            columns_in_datasource,
                        ]
                    )  # [file,most_recent_access,created,size]
                if include_stats == True and include_colnames == False:
                    file_list.append(
                        [
                            currentfolder,
                            time.ctime(folderstats.st_mtime),
                            time.ctime(folderstats.st_ctime),
                            file_size,
                            file_num,
                        ]
                    )  # [file,most_recent_access,created,size]
                if include_stats == False and include_colnames == False:
                    file_list.append(
                        [
                            currentfolder,
                            time.ctime(folderstats.st_mtime),
                            time.ctime(folderstats.st_ctime),
                        ]
                    )

    return file_list


def columnreplaceforsqlload(dataframe):
    dataframe.columns = dataframe.columns.str.strip()
    dataframe.columns = dataframe.columns.str.replace(" ", "_")
    dataframe.columns = dataframe.columns.str.replace("[()-.,:]", "_", regex=True)
    dataframe.columns = dataframe.columns.str.replace("[", "", regex=False)
    dataframe.columns = dataframe.columns.str.replace("]", "", regex=False)
    dataframe.columns = dataframe.columns.str.lower()
    

    if "grant" in dataframe.columns:
        dataframe.rename(columns={"grant": "_grant"}, inplace=True)
    return dataframe


# once we read in the data without a header this takes whatever is teh first full non_na row and sets that row as the header and resets the index once done
def skiprowstofirstfullnonna(dataframe):
    # note this code is cause concur files sometimes have useless crap at the bottom and we need to remove it as well as remove the column
    if len(dataframe.columns) == 3:
        dataframe = dataframe.iloc[:-1].dropna(axis=1, how="all")
    first_valid_loc = dataframe.apply(lambda col: col.first_valid_index()).max()
    dataframe.columns = dataframe.loc[first_valid_loc]
    # set the file to read in without headers, then get first full row, make that the headers, then remove all rows above and including that row
    dataframe = dataframe.loc[(first_valid_loc + 1) :]
    return dataframe


def path_leaf(path):
    # finds the file name from a path
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)


def create_csv_from_html(path:str, folder:str, final_folder:str ="none", since_in_seconds: float|None= None):
    '''
    This creates CSV from HTML (or rather xls, which is how Ariba exports them) and deletes the
    old xls file.

    Path - a path either complete or partial
    Folder-the folder name that we are using to search in
    final_folder- if we need to move the files to a new folder
    since_in_seconds- a timer to say only take the *most recent* file, can be done and then we do all files but don't
    delete
    '''
    if final_folder == "none":
        final_folder = folder
    # Changes directory to the folder where the files are in
    os.chdir(path + folder)
    # gets the most recent file from the folder
    file_failures = []
    return_value = ''
    if since_in_seconds == None:
        list_of_files = glob.glob(os.getcwd() + "\*.xls")  # * means all if need specific
        os.chdir(path + final_folder)
        for i in list_of_files:
            try:
                filename = path_leaf(i)
                csv_file = pd.read_html(i)
                csv_file[1].to_csv(filename.replace(".xls", "").replace(" ", "_") + ".csv")
            except:
                file_failures = file_failures.append(i)

    else:
        all_files = pd.DataFrame(
            get_information(os.getcwd(), only_files=True, file_end='*.xls'),
            columns=["name", "modified", "created"],
        )
        try:
            all_files["created"] = pd.to_datetime(
                all_files["created"], format="%a %B %d %H:%M:%S %Y"
            )
        except ValueError:
            all_files["created"] = pd.to_datetime(
                all_files["created"], format="%a %b %d %H:%M:%S %Y"
            )
        all_files["deltafromnow"] = (
            datetime.datetime.now() - all_files["created"]
        ) # type: ignore
        all_files_1day = all_files.loc[all_files["deltafromnow"]
            <= (datetime.timedelta(seconds=since_in_seconds))]
        all_file_list = all_files_1day['name'].tolist()
        if len(all_file_list) >= 1:
            os.chdir(path + final_folder)
            for i in all_file_list:
                try:
                    replace_path = str.replace(path+folder,'/','\\')
                    filename =  str.replace(i,replace_path+'\\','')
                    csv_file = pd.read_html(i)
                    csv_file[1].to_csv(filename.replace(".xls", "").replace(" ", "_") + ".csv")
                    if 'Monthly Contract' not in i:
                        os.remove(i)
                except:
                    file_failures.append(i)
        else:
            pass
    if file_failures==[]:
        return_value = 'All files succeeded'
    else:
        return_value = file_failures
    return return_value

        

def tkinter_directory():
    # if we are ever going to update to a tkinter user interface we would use this
    global folder_selected
    folder_selected = filedialog.askdirectory(initialdir=r"C:\\", title="Folder Select")
    pass


def error_function():
    "returns the error as a string so that we can act on it"
    # Get current system exception
    ex_type, ex_value, ex_traceback = sys.exc_info()

    # Extract unformatter stack traces as tuples
    trace_back = traceback.extract_tb(ex_traceback)

    # Format stacktrace
    stack_trace = list()
    pass

    for trace in trace_back:
        stack_trace.append(
            "File : %s , Line : %d, Func.Name : %s, Message : %s"
            % (trace[0], trace[1], trace[2], trace[3])
        )

    ex_type_append = "Exception type : %s " % ex_type.__name__
    ex_mes_append = "Exception message : %s" % ex_value
    ex_trace_append = "Stack trace : %s" % stack_trace
    ex_final_message = ex_type_append + " " + ex_mes_append + " " + ex_trace_append
    return ex_final_message


def convertcolumnswithdateinname(dataframe):
    """
    finds those columns with dates in the name and converts them to a datetype in pandas, coercing errors"
    """
    dataframe[[c for c in dataframe if (c.find("date") != -1)]] = dataframe[
        [c for c in dataframe if (c.find("date") != -1)]
    ].apply(pd.to_datetime, errors="coerce")
    return dataframe


def continue_from_sql_index(
    schemaandtableastext: str,
    loadingdataframe: object,
    engineinfunction: object,
    asindex: bool = True,
):
    """
    Grabs the maxindex from SQL and then adds it to the dataframe so that we continue to count from there

    This function requires a loading data frame, and a sql alchemy engine. It should almost always be included in a try except block
    to ensure if the table does not exist it restarts the count

    schemandtablestext - should include both the schema and table as text (for example dbo.table)
    loadingdataframe - the dataframe you're loading
    engineinfunction - the sqlalchemy engine you're using to connect
    """
    with engineinfunction.connect() as conn:
        result = conn.execute(text("SELECT Max([index]) FROM " + schemaandtableastext))
        for row in result:
            maxindex = row[0]
    if maxindex is None:
        maxindex = -1
    loadingdataframe["index"] = range(
        maxindex + 1, maxindex + 1 + len(loadingdataframe)
    )
    if asindex == True:
        loadingdataframe.set_index("index", inplace=True)
    return loadingdataframe


def readmostrecentexcel(filepath):
    mostrecentexcel = glob.glob(filepath)
    dataframe = pd.read_excel(max(mostrecentexcel, key=os.path.getctime), dtype=str)
    return dataframe


def find_files_created_today(directory: str):
    try:
        today_date = datet.now()
        filedataframe = pd.DataFrame(
            get_information(directory, only_files=True),
            columns=["name", "modified", "created"],
        )
        filedataframe["created"] = pd.to_datetime(filedataframe["created"])
        filescreatedtoday = filedataframe.loc[
            today_date - td(hours=24) <= filedataframe["created"]
        ]
    except:
        filescreatedtoday = pd.DataFrame()
    return filescreatedtoday


def LoadCsvsFromFolder(filelist, emptydataframe=pd.DataFrame()):
    for file in filelist:
        interim = pd.read_csv(file, dtype=str)
        emptydataframe = pd.concat([interim, emptydataframe], axis=0, join="outer")
    return emptydataframe


def LoadFilesFromFolder(filelist: list, filetype: str, emptydataframe=pd.DataFrame()):
    """
    This is an extrmely simplistic load of all files from a folder.
    It loads all file found in a glob list (not included as the regex should be obvious,
    and you may need it elsewhere) and places them in a brand new dataframe.
    It takes nothing other than csv, excel and html files.
    """
    filetypes = ["csv", "html", "excel"]
    if filetype not in filetypes:
        raise ValueError(
            "Invalid File Type. Only csv, excel and html files can be loaded. Note, this expects these csv, html or excel files to be perfect."
        )
    if filetype.lower() == "csv":
        for file in filelist:
            interim = pd.read_csv(file, dtype=str)
            emptydataframe = pd.concat([interim, emptydataframe], axis=0, join="outer")
    if filetype.lower() == "excel":
        for file in filelist:
            interim = pd.read_excel(file, dtype=str)
            emptydataframe = pd.concat([interim, emptydataframe], axis=0, join="outer")
    if filetype.lower() == "html":
        for file in filelist:
            interim = pd.read_html(file)
            interimdf = interim[1]
            emptydataframe = pd.concat(
                [interimdf, emptydataframe], axis=0, join="outer"
            )
    return emptydataframe

def test_created(filename, last_time):
    """
    This tests for the most recent file, and is built to be used in list comprehension
    
    """
    create_time = datetime.datetime.fromtimestamp(os.path.getctime(filename))
    last_time = pd.to_datetime(last_time)
    if create_time >= last_time:
        return True
    return False


def LoadFilesFromFolderSwitchable(
    regexfilepath: str,
    filetype: str,
    filestoload: str = "all",
    includefilenamecolumn=False,
    emptydataframe=pd.DataFrame(),
    numrowstoskip: int = 0,
    sheet_number: int = 0,
    file_encoding: str | None = None,
    files_since  = None):
    """
    This is a simple load of all, the most recent files from a folder, or since certain time, handled by choosing all, or recent in the filestoload
    It loads all file found in the file path whcih needs to be a regex formulation.
    files_since is included it loads them based on a certain time, and only
    For example, pathtofile/excelfiles*.xlsx
    The thing than places the file(s) in a brand new dataframe with the datatypes as string.
    It takes nothing other than csv, excel, fwf and html files.
    additional parameters are:

    filestoload = the file to load, you can enter all, recent, or .

    includefilenameincolumn = do you want the filename as an additional column, default is false
    
    emptydateframe = you can append if necessary to a data frame, so you could load
    the same file from multiple sources.
    
    numrowstoskip = allows for skipping multipel rows
    
    sheetnubmer = for excel, you can specify the sheet number as integer

    files_since = a Timestamp object that can be used to find files since x
    """
    filetypes = ["csv", "html", "excel", "fwf"]
    filestoloadoptions = ["all", "recent"]
    if filetype.lower() not in filetypes:
        raise ValueError(
            "Invalid File Type. Only csv, excel, fwf, and html files can be loaded. Note, this expects these csv, html or excel files to be perfect."
        )
    if filestoload.lower() not in filestoloadoptions:
        raise ValueError("Please choose all or most recent, will default to all")
    filelist = glob.glob(regexfilepath)
    if files_since is not None:
        filelist = [filename for filename in filelist if test_created(filename,files_since)]
    mostrecentfile = max(filelist, key=os.path.getctime)

    if filetype.lower() == "csv":
        if filestoload.lower() == "all":
            for file in filelist:
                interim = pd.read_csv(file, skiprows=numrowstoskip, dtype=str, encoding = file_encoding)
                if includefilenamecolumn == True:
                    interim["filename"] = file
                emptydataframe = pd.concat(
                    [interim, emptydataframe], axis=0, join="outer"
                )
            emptydataframe.reset_index(drop=True, inplace=True)
        else:
            emptydataframe = pd.read_csv(
                mostrecentfile, skiprows=numrowstoskip, dtype=str, encoding = file_encoding
            )
    if filetype.lower() == "excel":
        if filestoload.lower() == "all":
            for file in filelist:
                interim = pd.read_excel(
                    file, sheet_name=sheet_number, skiprows=numrowstoskip, dtype=str
                )
                if includefilenamecolumn == True:
                    interim["filename"] = file
                emptydataframe = pd.concat(
                    [interim, emptydataframe], axis=0, join="outer"
                )
            emptydataframe.reset_index(drop=True, inplace=True)
        else:
            emptydataframe = pd.read_excel(
                mostrecentfile,
                sheet_name=sheet_number,
                skiprows=numrowstoskip,
                dtype=str,
            )
    if filetype.lower() == "html":
        if filestoload.lower() == "all":
            for file in filelist:
                interim = pd.read_html(file)
                interimdf = interim[1]
                if includefilenamecolumn == True:
                    interimdf["filename"] = file
                emptydataframe = pd.concat(
                    [interimdf, emptydataframe], axis=0, join="outer"
                )
            emptydataframe.reset_index(drop=True, inplace=True)
        else:
            interim = pd.read_html(max(filelist, key=os.path.getctime))
            emptydataframe = interim[1]
    if filetype.lower() == "fwf":
        if filestoload.lower() == "all":
            for file in filelist:
                interim = pd.read_fwf(file, skiprows=numrowstoskip, dtype=str)
                if includefilenamecolumn == True:
                    interim["filename"] = file
                emptydataframe = pd.concat(
                    [interim, emptydataframe], axis=0, join="outer"
                )
            emptydataframe.reset_index(drop=True, inplace=True)
        else:
            emptydataframe = pd.read_fwf(
                mostrecentfile, skiprows=numrowstoskip, dtype=str
            )
    return emptydataframe


def update_sql_table(
    dataframe,
    sqltable,
    function_engine,
    function_connection,
    function_meta,
    date_columns=[],
    drop_duplicates=False,
):
    if sqltable.find("[") == -1:
        datatable = function_meta.tables[sqltable]
    else:
        datatable = function_meta.tables[sqltable.replace("[", "").replace("]", "")]
    pytypedict = {}
    sqltypedict = {}
    pycolumns = []
    
    for c in datatable.columns:
        pytypedict[c.name] = c.type.python_type.__name__
        sqltypedict[c.name] = c.type
        pycolumns += [c.name]

    dataframe.columns = pycolumns
    if drop_duplicates == True:
        dataframe.drop_duplicates(inplace=True)
    if date_columns != []:
        pytypedict = {
            k: ("datetime64[ns]" if k in date_columns else v)
            for k, v in pytypedict.items()
        }
        dataframe.astype(pytypedict)
    delete_stmt = text("DELETE FROM " + sqltable)
    function_connection.execute(delete_stmt)
    function_connection.commit()
    print("readytosql")
    try:
        dataframe.to_sql(
            sqltable,
            function_engine,
            if_exists="append",
            index=False,
            chunksize=100,
            dtype=sqltypedict,
        )
    except:
        try:
            dataframe.to_sql(
                sqltable,
                function_engine,
                if_exists="append",
                index=False,
                chunksize=100,
            )
        except:
            dataframe.astype(pytypedict)
            dataframe.to_sql(
                sqltable,
                function_engine,
                if_exists="append",
                index=False,
                chunksize=100,
                dtype=sqltypedict,
            )
    print("donewithsqlload")
    rebuild_stmt = text("ALTER INDEX ALL ON " + sqltable + " REBUILD")
    function_connection.execute(rebuild_stmt)
    function_connection.commit()
    print("donewithindexrebuild")
