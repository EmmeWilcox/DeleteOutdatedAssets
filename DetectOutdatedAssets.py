##Written by Grayson Wilcox
##wilcox.g8@gmail.com
##7-7-2016

#!/usr/bin/python

########################################################
##BEFORE USE, pip install ldap and pip install pymssql##
## Ensure that your user is allowed to use uslsasql1  ##
########################################################



##README##
##If you want to update location of a server, edit one of the following:
LDAP_SERVER = 'xxxxxxxxxxxxxxxxxx'
SQL_SERVER = 'xxxxxxxxxxxxxxxxxxxxxxx'
SQL_TABLE = 'servicedesk'
##If you need to update the table, make sure to check fetchAssetItems()

import ldap
import pymssql
import sys
import getpass
import csv
import os
import datetime

##Goal: To generate a list of active directory users.
##Params: Address is automatic but can be overwritten. It is the LDAP location of the AD server
##      : Username is sent in by the user. Requires ADMIN
##      : Password is obviously the username's password
##Return value: A list of all usernames trimmed of their email tag.
def fetchDisabledUsers(address, username, password):
    conn = ldap.initialize('ldap://' + address)
    conn.protocol_version = 3
    conn.set_option(ldap.OPT_REFERRALS, 0)
    l = conn
    try:
        l.simple_bind_s(username,password)
    except:
        print "Invalid credentials"
        sys.exit()

    ##Dig to Disabled users group.
    baseDN = "ou=Disabled,ou=Users,ou=xxxxxxx,dc=xxxxxxxxxxxxxxxxxx,dc=local"
    searchScope = ldap.SCOPE_SUBTREE
    retrieveAttributes = None
    ##Honestly not sure the value, should return all disabled.
    searchFilter = "(UserAccountControl:1.2.840.113556.1.4.803:=2)"
    
    try:
        ldap_result_id = l.search(baseDN, searchScope, searchFilter, retrieveAttributes)
        result_set = []
        while 1:
            result_type, result_data = l.result(ldap_result_id, 0)
            if (result_data == []):
                break
            else:
                if result_type == ldap.RES_SEARCH_ENTRY:
                    try:
                        me = result_data[0][1]['userPrincipalName'][0]
                        result_set.append(me[0:findEmail(me)])
                    except:
                        ##For documentation purposes, uncomment these. They will never be human accounts.
                        ##print result_data
                        ##print '\n\n\n'
                        if result_data == None:
                            break;
        ##print result_set
        return result_set
    except ldap.LDAPError, e:
        print 'No connection to Active Directory'
        sys.exit()

        

##Goal: To reach into the servicedesk resource server and retrieve all assets
##Params: Server is the address of the server, shouldn't be changed.
##      : Username is the user's name obviously
##      : Password is the user's password
##      : Database is servicedesk. This can be changed if we get a different service.
##Return Value: A list of all user owned (probably) devices
def fetchAssetItems(server,username,password,database):
    try: 
        db = pymssql.connect(server,username,password,database)
        cursor = db.cursor()
            
        ##Edit the following line if changing the table
        cursor.execute("select [RESOURCENAME],[SERIALNO],[ASSETTAG] from [servicedesk].[dbo].[Resources]")
        data = (cursor.fetchall())
        realdata = []
   
        for tup in data:
            retup = tup
            
            ##Encodes as ascii, it normally recieves it as a unicode object
            temp = retup[0].encode('ascii')
            
            ##Checks to ensure SERIALNO is not None
            if retup[1]:
                tempSer = ''.join(retup[1]).encode('ascii')
            else:
                tempSer = 'NA'

            ##Checks to ensure ASSETTAG is not None
            if retup[2]:
                tempAsset = ''.join(retup[2]).encode('ascii')
            else:
                tempAsset = 'NA'

            ##Makes all 3 a dictionary for easy packaging
            entry = {'Name': temp, 'Serno': tempSer, 'Asset': tempAsset}
            realdata.append(entry)
            
        db.close()
        
            
        return realdata
    except Exception as e:
        print e
        print 'You do not have permissions to access SQL database'
        sys.exit()

##Goal: To return all users who still have assets in the servicedesk resources database
##Params: AD is the list returned by fetchDisabledUsers
##      : Assets is the list returned by fetchAssetItems
##Return Value: A list of all items to be deleted. 
def compareAssets(AD,Assets):
    overdrawnAssets = []
    for person in AD:
        for asset in Assets:
            if person.lower() in asset['Name'].lower():
                overdrawnAssets.append(asset)
    return overdrawnAssets
    
        


def compareDuplicates(server,username,password,database,serials):
    try: 
        db = pymssql.connect(server,username,password,database)
        cursor = db.cursor()
        ##Edit the following line if changing the table
        cursor.execute("select [RESOURCENAME],[SERIALNO] from [servicedesk].[dbo].[Resources]")
        SerName = (cursor.fetchall())
        data = []
        for foo in serials:
            if foo['Serno'] == 'NA':
                continue
            mine = (foo['Serno'],)
            for tup in SerName:
                if foo['Serno'] == SerName[1]:
                    mine+=SerName[0]
            if len(mine) - 1:
                data.append(mine)
        db.close()
        return data
    
    except Exception as e:
        print "Something went wrong..."
        sys.exit()
        


##Finds the index of an @. 
def findEmail(shoestring):
    for i, n in enumerate(shoestring):
        if n != '@':
            continue
        else:
            return i


##Main
user = raw_input('Input an admin username.(Do not add domain)  ')
##Getpass will only work on cmd.
pswd = getpass.getpass('Input your password.  ')
##Sets domain (put your domain for xxxxx
user = """xxxxxx""" + user

print 'Fetching Disabled Users'    
ADList = fetchDisabledUsers(LDAP_SERVER,user,pswd)

print 'Fetching All Applicable Assets'
AssetList = fetchAssetItems(SQL_SERVER,user,pswd,SQL_TABLE)


print 'Comparing All Assets'
badAssets = compareAssets(ADList,AssetList)

##Check for having no issues
if not badAssets:
    print 'No issues detected!'
    sys.exit()

print 'Checking for duplicates'
DuplicateList = compareDuplicates(SQL_SERVER,user,pswd,SQL_TABLE,badAssets)

##Make sure there are some duplicates.
shouldOpenDuplicates = bool(len(DuplicateList))

##This tells the program where it is.
selfPath = os.path.abspath('')

print 'Making csv'
##Shows how many issues there are.
i = 0
csvname = 'outdated_assets_' + str(datetime.date.today()) + '.csv'
with open (selfPath + '\\' + csvname,'wb') as csvfile:
    iWrite = csv.writer(csvfile, delimiter=',')
    iWrite.writerow(["There are " + str(len(badAssets)) + " issues."])
    
    for asset in badAssets:
        iWrite.writerow([asset['Name'],asset['Serno'],asset['Asset']])
        i+=1


##Shows how many re-entries there are
if shouldOpenDuplicates:
    csvname2 = 'duplicate_assets_' + str(datetime.date.today()) + '.csv'
    with open (selfPath + '\\' + csvname2,'wb') as csvfile:
        iWrite = csv.writer(csvfile2, delimiter=',')
        iWrite.writerow(["There are " + str(len(DuplicateList)) + " duplicates."])
        for item in DuplicateList:
            shoestring = ''
            for it in item:
                shoestring+=it + ','
            iWrite.writerow(shoestring)

##Opens the csv files literally
try:
    os.startfile(selfPath + "\\"+csvname)
except:
    print "Failed to launch document, please close any open copies"

try:
    os.startfile(selfPath + "\\"+csvname2)
except:
    print "No duplicates"
