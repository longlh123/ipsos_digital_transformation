import os
import sys
sys.path.append(r'C:\Users\long.pham\Documents\MDDPython\ipsos_digital_transformation\libs')
import re
import numpy as np
import pandas as pd
import win32com.client as w32
import shutil
from libs.metadata import Metadata
from datetime import datetime 

import collections.abc
#hyper needs the four following aliases to be done manually.
collections.Iterable = collections.abc.Iterable
collections.Mapping = collections.abc.Mapping
collections.MutableSet = collections.abc.MutableSet
collections.MutableMapping = collections.abc.MutableMapping

os.chdir("ipsos_digital_transformation")

clean = re.compile('<.*?>')

#dien thong tin ten du an
project_name = "VN2023200DIGITAL_THEGLORY"

root = r"projects\{}".format(project_name)
excel_path = r"{}\{}".format(root, "out2023-09-17.xlsx")

source_dms_file =  r"dms\OutputDDFFile.dms"
source_mdd_file =  r"template\TemplateProject.mdd"
current_mdd_file = r"{}\{}".format(root, "{}.mdd".format(project_name))

if not os.path.exists(current_mdd_file):
    shutil.copy(source_mdd_file, current_mdd_file)

mdd_source = Metadata(mdd_file=current_mdd_file, dms_file=source_dms_file)

questions = dict()

df_datasource = pd.read_excel(excel_path, engine="openpyxl", sheet_name="Labels", header=[0,1])
df_datasource_codes = pd.read_excel(excel_path, engine="openpyxl", sheet_name="Codes", header=[0,1])

main_columns = list(df_datasource.columns)

i = 0

while i < len(main_columns):
    step = 1
    c = main_columns[i]

    if "IDP205" in c[0]:
        a = ""
    
    if c[0] not in ["IDP207","IDP208","IDP209"]:
        if re.match(pattern="^(\w+)\s(\(.+\))__(\d+)$", string=c[0]):
            #GRID (MA)
            m = re.search(pattern="^(\w+)", string=c[0])

            question_name = "{}_LOOP".format(c[0][m.span()[0] : m.span()[1]])

            if question_name not in list(questions.keys()):
                questions[question_name] = dict()

            questions[question_name]["question_name"] = question_name
            
            if "attributes" not in questions[question_name].keys():
                questions[question_name]["attributes"] = dict()
            
            attribute_name = "R{}".format(len(questions[question_name]["attributes"].keys()) + 1)

            attribute_column = re.sub(pattern="__(\d+)$", repl="", string=c[0])
            attribute_column = attribute_column.replace("?", "\?").replace("(", "\(").replace(")", "\)")

            attribute_columns = [col for col in list(df_datasource.columns) if re.match(pattern="^({})__(\d+)$".format(attribute_column), string=col[0])]

            if "categories" not in questions[question_name].keys():
                questions[question_name]["categories"] = dict()

            for attr_column in attribute_columns:
                attr_match = re.search(pattern="^[^-]+", string=c[1])
                
                questions[question_name]["attributes"][attribute_name] = attr_column[1][attr_match.span()[0] : attr_match.span()[1]]
                
                category_name = re.sub(pattern="^({})__".format(attribute_column), repl="_", string=attr_column[0])

                if category_name not in questions[question_name]["categories"].keys():
                    questions[question_name]["categories"][category_name] = dict()
                    questions[question_name]["categories_inplace"] = dict()
                
                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern="^(.+)-", repl="", string=attr_column[1]).strip()

                questions[question_name]["categories_inplace"]["Yes"] = category_name.replace("_", "")
                questions[question_name]["categories_inplace"]["No"] = np.nan

                df_datasource[attr_column].replace(questions[question_name]["categories_inplace"], inplace=True)
            
            df_datasource["%s[{%s}]._Codes" % (question_name, attribute_name)] = df_datasource[attribute_columns].apply(lambda x: np.nan if x.count() == 0 else str("{") + ','.join(["_{}".format(str(i)) for i in list(x) if pd.isna(i) is False]) + str("}"), axis=1)
            df_datasource.drop(columns=attribute_columns, inplace=True)
            
            step = len(attribute_columns)

            m = re.search(pattern="^(\w+)", string=main_columns[i + step][0])

            question_name_next = "{}_LOOP".format(main_columns[i + step][0][m.span()[0] : m.span()[1]])

            if question_name != question_name_next:
                question_syntax = '%s "%s"\nloop{\n' % (question_name, question_name)

                for aid, attribute in questions[question_name]["attributes"].items():
                    attribute_syntax = '\t%s "%s"' % (aid, attribute)

                    attribute_syntax += ",\n" if list(questions[question_name]["attributes"].keys()).index(aid) < len(list(questions[question_name]["attributes"].keys())) - 1 else "\n"

                    question_syntax += attribute_syntax 

                question_syntax += '}fields()expand grid;\n\n'
                
                questionc_child_syntax = '_Codes "Codes" categorical[1..]\n{\n'

                for cid, category in questions[question_name]["categories"].items():
                    category_syntax = '\t%s "%s"' % (cid, category["label"])

                    category_syntax += ",\n" if list(questions[question_name]["categories"].keys()).index(cid) < len(list(questions[question_name]["categories"].keys())) - 1 else "\n"

                    questionc_child_syntax += category_syntax 

                questionc_child_syntax += '};\n\n'

                mdd_source.addScript(questions[question_name]["question_name"], question_syntax, childnodes=[questionc_child_syntax])
        elif re.match(pattern="^(\w+)\s(\(.+\))$", string=c[0]):
            #GRID (SA)
            m = re.search(pattern="^(\w+)", string=c[0])

            question_name = "{}_LOOP".format(c[0][m.span()[0] : m.span()[1]])

            if question_name not in list(questions.keys()):
                questions[question_name] = dict()

            questions[question_name]["question_name"] = question_name
            questions[question_name]["question_text"] = question_name

            if "attributes" not in questions[question_name].keys():
                questions[question_name]["attributes"] = dict()

            attribute_columns = [col for col in list(df_datasource.columns) if re.match(pattern="^({})\s(\(.+\))$".format(c[0][m.span()[0] : m.span()[1]]), string=col[0])]

            for attr_column in attribute_columns:
                attribute_name = "R{}".format(len(questions[question_name]["attributes"].keys()) + 1)

                questions[question_name]["attributes"][attribute_name] = attr_column[1]

                if all([s not in c[0] for s in ["IDP205","IDP206"]]):
                    df_datasource_codes[attr_column] = df_datasource_codes[attr_column].fillna(0).astype(np.int64)

                    df_datasource.loc[df_datasource[attr_column].notnull(), attr_column] = df_datasource_codes.loc[df_datasource_codes[attr_column].notnull(), attr_column].astype(str) + ". " + df_datasource.loc[df_datasource[attr_column].notnull(), attr_column].astype(str)

                    if "categories" not in questions[question_name].keys():
                        questions[question_name]["categories"] = dict()
                        questions[question_name]["categories_inplace"] = dict()

                    df_datasource[attr_column] = df_datasource[attr_column].astype(str)
                    categories = df_datasource.groupby([attr_column], axis=0).groups.keys()

                    idx_cat = 1
                    
                    for category in list(categories):
                        if category not in ["--", "nan", np.nan]:
                            if re.match(pattern="^([0-9]*)\.", string=str(category)):
                                cat_match = re.match(pattern="^([0-9]*)\.", string=category)
                                category_name = "_{0}".format(category[cat_match.span()[0]:cat_match.span()[1] - 1])
                                
                                if category_name not in questions[question_name]["categories"].keys():
                                    questions[question_name]["categories"][category_name] = dict()
                                    questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                
                                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                            else:
                                category_name = "_{0}".format(str(idx_cat))

                                if category_name not in questions[question_name]["categories"].keys():
                                    questions[question_name]["categories"][category_name] = dict()
                                    questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                
                                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                                    
                                idx_cat += 1     

                    df_datasource[attr_column].replace(questions[question_name]["categories_inplace"], inplace=True)
                    df_datasource.rename(columns={ attr_column[0] : "%s[{%s}]._Codes" % (question_name, attribute_name) }, inplace=True)
                else:
                    df_datasource.rename(columns={ attr_column[0] : "%s[{%s}]._Text" % (question_name, attribute_name) }, inplace=True)

            question_syntax = '%s "%s"\nloop{\n' % (question_name, questions[question_name]["question_text"])

            for aid, attribute in questions[question_name]["attributes"].items():
                attribute_syntax = '\t%s "%s"' % (aid, attribute)

                attribute_syntax += ",\n" if list(questions[question_name]["attributes"].keys()).index(aid) < len(list(questions[question_name]["attributes"].keys())) - 1 else "\n"

                question_syntax += attribute_syntax 
            
            if all([s not in c[0] for s in ["IDP205","IDP206"]]):
                question_syntax += '}fields(_Codes "Codes" categorical[1..1]\n{\n'

                for cid, category in questions[question_name]["categories"].items():
                    category_syntax = '\t%s "%s"' % (cid, category["label"])

                    category_syntax += ",\n" if list(questions[question_name]["categories"].keys()).index(cid) < len(list(questions[question_name]["categories"].keys())) - 1 else "\n"

                    question_syntax += category_syntax 

                question_syntax += '})expand grid;\n\n'
            else:
                question_syntax += '}fields(_Text "Text" text;)expand grid;\n\n'
            
            mdd_source.addScript(questions[question_name]["question_name"], question_syntax)

            step = len(question_columns)

        elif re.match(pattern="^(\w+)__(\d+)$", string=c[0]):
            #CATEGORICAL (MA)
            question_name = c[0].split("__")[0]

            if question_name not in list(questions.keys()):
                questions[question_name] = dict()
            
            questions[question_name]["question_name"] = question_name

            if "categories" not in questions[question_name].keys():
                questions[question_name]["categories"] = dict()

            question_columns = [col for col in list(df_datasource.columns) if re.match(pattern="^({})__(\d+)$".format(question_name), string=col[0])]
            
            for q_column in question_columns:
                qre_match = re.match(pattern="(.+)\?", string=c[1])
                
                questions[question_name]["question_text"] = q_column[1].replace("\n", "") if qre_match is None else q_column[1][qre_match.span()[0]:qre_match.span()[1]]
                
                category_name = re.sub(pattern="^(\w+)__", repl="_", string=q_column[0])

                if category_name not in questions[question_name]["categories"].keys():
                    questions[question_name]["categories"][category_name] = dict()
                    questions[question_name]["categories_inplace"] = dict()
                
                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern="^(.+)-", repl="", string=q_column[1]).strip()

                questions[question_name]["categories_inplace"]["Yes"] = category_name.replace("_", "")
                questions[question_name]["categories_inplace"]["No"] = np.nan

                df_datasource[q_column].replace(questions[question_name]["categories_inplace"], inplace=True)

            df_datasource[question_name] = df_datasource[question_columns].apply(lambda x: np.nan if x.count() == 0 else str("{") + ','.join(["_{}".format(str(i)) for i in list(x) if pd.isna(i) is False]) + str("}"), axis=1)
            df_datasource.drop(columns=question_columns, inplace=True)

            question_syntax = '%s "%s"\ncategorical%s\n{\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "[1..]")

            for cid, category in questions[question_name]["categories"].items():
                category_syntax = '\t%s "%s"' % (cid, category["label"])

                category_syntax += ",\n" if list(questions[question_name]["categories"].keys()).index(cid) < len(list(questions[question_name]["categories"].keys())) - 1 else "\n"

                question_syntax += category_syntax 

            question_syntax += '};\n\n'
            
            mdd_source.addScript(questions[question_name]["question_name"], question_syntax)

            step = len(question_columns)

            if i + step < len(main_columns):
                c_next = main_columns[i + step]

                if re.match(pattern="^({})_([A-Za-z0-9]+)$".format(question_name), string=c_next[0]):
                    question_name = c_next[0]

                    if question_name not in list(questions.keys()):
                        questions[question_name] = dict()
                    
                    questions[question_name]["question_name"] = question_name

                    qre_match = re.match(pattern="(.+)\?", string=c[1])

                    questions[question_name]["question_text"] = c_next[1] if qre_match is None else c_next[1][qre_match.span()[0] : qre_match.span()[1]]

                    df_datasource.rename(columns={ c_next[0] : questions[question_name]["question_name"] }, inplace=True)

                    mdd_source.addScript(questions[question_name]["question_name"], '%s "%s" %s;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "text"))

                    step += 1
        else:
            is_grid_sa = False
            
            if i + 1 < len(main_columns):
                c_next = main_columns[i + 1]

                if re.match(pattern="^(CHILD)_{}$".format(c[0]), string=main_columns[i + 1][0]):
                    #GRID (SA)
                    is_grid_sa = True
                    question_name = c[0]

                    if question_name not in list(questions.keys()):
                        questions[question_name] = dict()

                    questions[question_name]["question_name"] = c[0]
                    questions[question_name]["question_text"] = c[1]

                    if "attributes" not in questions[question_name].keys():
                        questions[question_name]["attributes"] = dict()

                    attribute_columns = [col for col in list(df_datasource.columns) if re.match(pattern="^({})\s(\(.+\))$".format(main_columns[i + 1][0]), string=col[0])]

                    for attr_column in attribute_columns:
                        attribute_name = "R{}".format(len(questions[question_name]["attributes"].keys()) + 1)

                        questions[question_name]["attributes"][attribute_name] = attr_column[1]

                        df_datasource_codes[attr_column] = df_datasource_codes[attr_column].fillna(0).astype(np.int64)

                        df_datasource.loc[df_datasource[attr_column].notnull(), attr_column] = df_datasource_codes.loc[df_datasource_codes[attr_column].notnull(), attr_column].astype(str) + ". " + df_datasource.loc[df_datasource[attr_column].notnull(), attr_column].astype(str)

                        if "categories" not in questions[question_name].keys():
                            questions[question_name]["categories"] = dict()
                            questions[question_name]["categories_inplace"] = dict()

                        df_datasource[attr_column] = df_datasource[attr_column].astype(str)
                        categories = df_datasource.groupby([attr_column], axis=0).groups.keys()

                        idx_cat = 1
                        
                        for category in list(categories):
                            if category not in ["--", "nan", np.nan]:
                                if re.match(pattern="^([0-9]*)\.", string=str(category)):
                                    cat_match = re.match(pattern="^([0-9]*)\.", string=category)
                                    category_name = "_{0}".format(category[cat_match.span()[0]:cat_match.span()[1] - 1])
                                    
                                    if category_name not in questions[question_name]["categories"].keys():
                                        questions[question_name]["categories"][category_name] = dict()
                                        questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                    
                                    questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                                else:
                                    category_name = "_{0}".format(str(idx_cat))

                                    if category_name not in questions[question_name]["categories"].keys():
                                        questions[question_name]["categories"][category_name] = dict()
                                        questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                    
                                    questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                                        
                                    idx_cat += 1     

                        df_datasource[attr_column].replace(questions[question_name]["categories_inplace"], inplace=True)
                        df_datasource.rename(columns={ attr_column[0] : "%s[{%s}]._Codes" % (question_name, attribute_name) }, inplace=True)

                    question_syntax = '%s "%s"\nloop{\n' % (question_name, questions[question_name]["question_text"])

                    for aid, attribute in questions[question_name]["attributes"].items():
                        attribute_syntax = '\t%s "%s"' % (aid, attribute)

                        attribute_syntax += ",\n" if list(questions[question_name]["attributes"].keys()).index(aid) < len(list(questions[question_name]["attributes"].keys())) - 1 else "\n"

                        question_syntax += attribute_syntax 

                    question_syntax += '}fields(_Codes "Codes" categorical[1..1]\n{\n'

                    for cid, category in questions[question_name]["categories"].items():
                        category_syntax = '\t%s "%s"' % (cid, category["label"])

                        category_syntax += ",\n" if list(questions[question_name]["categories"].keys()).index(cid) < len(list(questions[question_name]["categories"].keys())) - 1 else "\n"

                        question_syntax += category_syntax 

                    question_syntax += '})expand grid;\n\n'
                    
                    mdd_source.addScript(questions[question_name]["question_name"], question_syntax)

                    step = len(question_columns) + 1
            
            if not is_grid_sa:
                question_name = re.sub(pattern="[-\s]", repl='_', string=c[0].strip())
                question_name = re.sub(pattern="[\(\)]", repl='', string=question_name)
                
                if question_name not in list(questions.keys()):
                    questions[question_name] = dict()

                questions[question_name]["question_name"] = question_name

                qre_match = re.match(pattern="(.+)\?", string=c[1])

                question_text = c[1] if qre_match is None else c[1][qre_match.span()[0] : qre_match.span()[1]]
                question_text = re.sub(pattern="\"", repl="\'", string=question_text)

                questions[question_name]["question_text"] = question_text
                
                if df_datasource[c].dtype.name not in ["object","str"]:
                    df_datasource.rename(columns={ c[0] : question_name }, inplace=True)

                    mdd_source.addScript(questions[question_name]["question_name"], '%s "%s" %s;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "double"))
                else:
                    df_datasource_codes[c] = df_datasource_codes[c].fillna(0).astype(np.int64)
                    
                    df_datasource.loc[df_datasource[c].notnull(), c] = df_datasource_codes.loc[df_datasource_codes[c].notnull(), c].astype(str) + ". " + df_datasource.loc[df_datasource[c].notnull(), c].astype(str) 

                    if "categories" not in questions[question_name].keys():
                        questions[question_name]["categories"] = dict()
                        questions[question_name]["categories_inplace"] = dict()

                    df_datasource[c] = df_datasource[c].astype(str)
                    categories = df_datasource.groupby([c], axis=0).groups.keys()

                    idx_cat = 1
                    
                    for category in list(categories):
                        if category not in ["--", "nan", np.nan]:
                            if re.match(pattern="^([0-9]*)\.", string=str(category)):
                                cat_match = re.match(pattern="^([0-9]*)\.", string=category)
                                category_name = "_{0}".format(category[cat_match.span()[0]:cat_match.span()[1] - 1])
                                
                                if category_name not in questions[question_name]["categories"].keys():
                                    questions[question_name]["categories"][category_name] = dict()
                                    questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                
                                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                            else:
                                category_name = "_{0}".format(str(idx_cat))

                                if category_name not in questions[question_name]["categories"].keys():
                                    questions[question_name]["categories"][category_name] = dict()
                                    questions[question_name]["categories_inplace"][category] = "{%s}" % (category_name)
                                
                                questions[question_name]["categories"][category_name]["label"] = re.sub(pattern=clean, repl="", string=str(category)) 
                                    
                                idx_cat += 1
                    
                    question_syntax = '%s "%s"\ncategorical%s\n{\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "[1..1]")

                    for cid, category in questions[question_name]["categories"].items():
                        category_syntax = '\t%s "%s"' % (cid, category["label"])
                        
                        category_syntax += ",\n" if list(questions[question_name]["categories"].keys()).index(cid) < len(list(questions[question_name]["categories"].keys())) - 1 else "\n"

                        question_syntax += category_syntax 

                    question_syntax += '};\n\n'

                    df_datasource[c].replace(questions[question_name]["categories_inplace"], inplace=True)
                    df_datasource.rename(columns={ c[0] : questions[question_name]["question_name"] }, inplace=True)
                    
                    mdd_source.addScript(questions[question_name]["question_name"], question_syntax)

                    if i + 1 < len(main_columns):
                        c_next = main_columns[i + 1]

                        if re.match(pattern="^({})_(\d+)$".format(question_name), string=c_next[0]):
                            questions[question_name]["question_name"] = c_next[0]

                            qre_match = re.match(pattern="(.+)\?", string=c[1])

                            questions[question_name]["question_text"] = c_next[1] if qre_match is None else c_next[1][qre_match.span()[0] : qre_match.span()[1]]

                            df_datasource.rename(columns={ c_next[0] : questions[question_name]["question_name"] }, inplace=True)

                            mdd_source.addScript(questions[question_name]["question_name"], '%s "%s" %s;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "text"))

                            step += 1
            
    else:
        question_name = re.sub(pattern=clean, repl="", string=c[0].strip().replace(" ", "_"))
        
        if question_name not in list(questions.keys()):
            questions[question_name] = dict()
        
        questions[question_name]["question_name"] = question_name
        
        qre_match = re.match(pattern="(.+)\?", string=c[1])

        questions[question_name]["question_text"] = c[1] if qre_match is None else c[1][qre_match.span()[0] : qre_match.span()[1]]
        
        df_datasource[c] = df_datasource[c].astype(str)

        if df_datasource[c].dtype.name in ["object","str"]:
            mdd_source.addScript(questions[question_name]["question_name"], '%s "%s" %s;\n\n' % (questions[question_name]["question_name"], questions[question_name]["question_text"], "text"))
    
    i += step

#df_datasource.drop(columns=["SCREENER11_13", "IDP178_IDPA684"], inplace=True)
df_datasource = df_datasource.droplevel(level=1, axis=1)

mdd_source.runDMS()
    
df_datasource.set_index(["Participant_Id"], inplace=True)

adoConn = w32.Dispatch('ADODB.Connection')
conn = "Provider=mrOleDB.Provider.2; Data Source = mrDataFileDsc; Location={}; Initial Catalog={}; Mode=ReadWrite; MR Init Category Names=1".format(mdd_source.mdd_file.replace('.mdd', '_EXPORT.ddf'), mdd_source.mdd_file.replace('.mdd', '_EXPORT.mdd'))
adoConn.Open(conn)

sql_delete = "DELETE FROM VDATA"
adoConn.Execute(sql_delete)

for i, row in df_datasource[list(df_datasource.columns)].iterrows():
    try:
        sql_insert = "INSERT INTO VDATA(Participant_Id) VALUES('%s')" % (i)
        adoConn.Execute(sql_insert)
        
        c = list()
        v = list()

        idx = 0

        for s in tuple(row):
            if not pd.isna(s) and s not in ['nan']:
                if row.index[idx] == "IDP178":
                    a = ""

                c.append(row.index[idx])

                if df_datasource[row.index[idx]].dtype.name in ["object","int"]:
                    if "datetime" in str(type(row[idx])):
                        s = s.strftime("%m/%d/%Y, %H:%M:%S")
                    else: 
                        s = s.replace("\n", "")

                    v.append("'{}'".format(s) if len(s) > 0 else np.nan)
                else:
                    v.append(s) 
            idx += 1    
        
        sql_update = "UPDATE VDATA SET " + ','.join([cx + str(r" = %s") for cx in c]) % tuple(v) + " WHERE Participant_Id = {}".format(int(i))
        adoConn.Execute(sql_update)    
    except Exception as ex:
        print(ex.excepinfo[1], ex, sep="-")
        sys.exit(1)
