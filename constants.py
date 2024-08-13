#
# Some departments should be ignored or renamed.  This dictionary contains those departments
# A blank value means 'ignore'; a non-blank value means 'rename'
#
DEPARTMENT_EXCEPTIONS = {"Learning Development": "", "VMT": "", "PSHCE": "", "Design & Technology":
                         "Design and Technology", "Information Technology": "Computer Science"}

#
# Spreadsheet names and path
#
REWARDS_SPREADSHEET = "spreadsheets/Senior Department Rewards and Sanctions.xlsx"
CONTACTS_SPREADSHEET = "spreadsheets/Pupil Contacts.xlsx"

#
# Database name and path
#
DATABASE_NAME = "databases/AY2023_2024.db"

#
# Table creation/DDL strings
#
CREATE_PUPIL_TABLE = "CREATE TABLE IF NOT EXISTS pupil (puid TEXT PRIMARY KEY NOT NULL, academicyear INTEGER, \
                      fname TEXT, sname TEXT)"

CREATE_DEPT_TABLE = "CREATE TABLE IF NOT EXISTS '{department}' (uid INTEGER PRIMARY KEY NOT NULL,  \
                     pupilid TEXT, cumulativepoints INTEGER, FOREIGN KEY(pupilid) REFERENCES pupil(puid))"

#
# Query, Insert, Update/DML strings
#
# In some of the following, {} are used to denote template text to be replaced by real values in the code
#
QUERY_PUPIL_TABLE = "SELECT * FROM pupil WHERE puid="
CHECK_PUPIL_TABLE = "SELECT COUNT(*) FROM pupil WHERE puid="
INSERT_PUPIL_ROW = "INSERT INTO pupil (puid, academicyear, fname, sname) \
                    VALUES (?, ?, ?, ?)"
SELECT_PUPIL = "SELECT puid FROM pupil WHERE fname = '{fname}' and \
                sname = '{sname}' and academicyear = {academicyear}"

QUERY_DEPT_TABLE = "SELECT cumulativepoints FROM '{department}' WHERE pupilid = "
UPDATE_DEPT_ROW = "UPDATE '{department}' SET cumulativepoints = {newpoints} WHERE pupilid = "
INSERT_DEPT_ROW = "INSERT INTO '{department}' (pupilid, cumulativepoints) \
                    VALUES (?, ?)"

QUERY_FIRST_PUPIL = "SELECT puid FROM pupil ORDER BY ROWID ASC LIMIT 1"
QUERY_ALL_PUPILS = "SELECT puid FROM pupil"

#
# These should be exported from iSAMS into a spreadsheet, but for now at least, hardcode
#
DEPARTMENT_HEADS = {"Computer Science": "Mr M R Gamble", "Mathematics": "Mrs K F Conroy"}

#
# Columns for the master sheet
#
PUPIL_NAME: int = 1
PUPIL_YEAR: int = 5
PUPIL_POINTS: int = 6

#
# Columns for the contacts sheet
#
CONTACT_PREFERRED_NAME: int = 2
CONTACT_SURNAME: int = 3
CONTACT_YEAR: int = 6
CONTACT_EMAIL: int = 4
