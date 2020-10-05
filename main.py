import pandas as pd

file = "SituazioneIntegration.xlsx"

sql1 = "INSERT INTO tasks (`title`, `description`, `date_creation`, `date_due`, `color_id`, `project_id`, `column_id`, `position`, `score`, `is_active`, `category_id`, `date_modification`, `creator_id`, `reference`, `date_started`, `swimlane_id`, `date_moved`) "
sql3 = "FROM tasks WHERE column_id='5'"

xl = pd.ExcelFile(file)

print("Fogli disponibili: ", xl.sheet_names)

df = pd.read_excel(xl, 'Pivot x tipo macc. DA FARE', header=None)
# numerazione colonne excel in base 0
# A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC
# 0 1 2 3 4 5 6 7 8 9 + 1 2 3 4 5 6 7 8 9 + 1 2 3 4 5 6 7 8 9

# print(df.iat[0, 0])
# print(df.iat[6, 10])
# print(df.iloc[3])
modulo = ""
print("df.index: ", df.index)

print("df.columns: ", df.columns)
for test in range(4, len(df.index)):
    if isinstance(df.iat[test, 2], str):
        stri = df.iat[test, 2].strip()
        is_stringa = True
    else:
        stri = df.iat[test, 2]
        is_stringa = False

    # verifico che stri non sia la stringa "(vuoto)" di excel
    if is_stringa and stri != "(vuoto)":
        modulo = int(stri) % 100  # "yeah!"
    else:
        modulo = "fuck!"

    if modulo == 0:
        is_father="PADRE"
    else:
        is_father=""

    print(test, ":::", type(df.iat[test, 2]), "---stri:", stri, "---is_stringa:", is_stringa, "---", is_father)  # df.iat[test, 3])

# For x = iRigaStart To iRighe ' - iRigaStart
# Range("D" & x).Select
# appo1 = Replace(ActiveCell.Value, vbLf, "")
# Range("C" & x).Select
# appo1 = appo1 & " - " & ActiveCell.Value
#
# Range("E" & x).Select
# appo2 = ActiveCell.Value
# Range("G" & x).Select
# appo2 = appo2 & " - " & Replace(ActiveCell.Value, "'", "_")
#
# Range("B" & x).Select
# appo3 = ActiveCell.Value
#
# sql2 = "SELECT '" & Trim(appo1) & "', '" & Trim(
#     appo2) & "', UNIX_TIMESTAMP(), '0','yellow', '1', '5', MAX(position)+1 AS position, '0', '1', '1', UNIX_TIMESTAMP()+60, '1', '" & Trim(
#     appo3) & "', '0', '1', UNIX_TIMESTAMP() "
#
# txtFile.WriteLine(sql1 & sql2 & sql3)
#
# Next x
