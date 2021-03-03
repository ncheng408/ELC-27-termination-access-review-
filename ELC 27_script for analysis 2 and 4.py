#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import numpy as np
import os
import argparse
from datetime import date, timedelta, datetime

# function that determines the org based on job title, department and location
def determineOrganization(job, dept, loc):
    # if dept has "restaurant", then org is Hospitality
    if 'Restaurant_' in dept or 'Restaurant' in loc or 'Ma(i)sonry' in loc:
        return 'Hospitality'
    
    # if dept has MFG, then org is Manufacturing
    if 'MFG' in dept:
        return 'Manufacturing'
    
    if 'Outlet' in job or 'Outlet_' in dept:
        return 'Outlet'
    
    if 'Gallery_' in dept or 'Gallery' in job or job == 'Design Consultant' or job == 'Associate Designer':
        return 'Gallery'
    
    if 'CSC' in dept:
        return 'Delight'
    
    if 'HDL' in dept:
        return 'HDC'
    
    if 'DC' in dept or 'Distribution Center' in job or 'Warehouse Sale' in dept:
        return 'DC'
    
    if '_CE' in dept:
        return 'Delight'
    
    if 'Home Office' in dept or loc == 'Home Office':
        return 'Home Office'
    
    return 'UNKNOWN'

# function that determines if leader based on job title
def determineIsLeader(job):
    return 'Leader' in job

# function that extracts the namd and assoc id based on line in BPM report
def extractTerminationNameAndId(subject):
    if not subject.startswith('Termination for'):
        return None, None
    
    trimmed = subject[len('Termination for '):]
    
    # name is split by comma
    parts = trimmed.split(',')
    
    if len(parts) < 2:
        return None, None
    
    name = parts[0]
    
    # assoc id is separated by a space
    moreparts = parts[1].strip().split(' ')
        
    if len(moreparts) < 2:
        return name, None
    
    return name, moreparts[0]

if __name__ == '__main__':
	
	## PLEASE UPDATE THESE #####################

	datapath = ['ELC 27','Q4 2020','Reports','Oracle HCM']

	# read in Levisay file
	levisayfile = datapath.copy()
	levisayfile.append('Termination Report Levisay.csv')
	termdata = pd.read_csv(os.path.join(*levisayfile))

	# read in BPM file
	bpmfile = datapath.copy()
	bpmfile.append('BPM Q4 02 24 21.csv')
	bpmdata = pd.read_csv(os.path.join(*bpmfile))

	# read in Last Login file
	addatapath = ['ELC 27','Q4 2020','Reports','AD']
	lastlogonfile = addatapath.copy()
	lastlogonfile.append("User's_Last_Logon_1612271987781_1.csv")
	logondata = pd.read_csv(os.path.join(*lastlogonfile), skiprows=8)

	# read in All Users file
	allusersfile = addatapath.copy()
	allusersfile += ['ADMPReport.csv']
	usersdata = pd.read_csv(os.path.join(*allusersfile))

	# read in Recently Disabled Users file
	disabledfile = addatapath.copy()
	disabledfile.append('Recently_Disabled_Users_1612272038965_1.csv')
	disableddata = pd.read_csv(os.path.join(*disabledfile), skiprows=8)
	
	# update Start and End of quarter dates
    	stdate = '2020-11-01'
    	eddate = '2021-01-30'
	
	##############################################

	## process levisay file
	# convert to datetime
	termdata['td'] = pd.to_datetime(termdata['Termination Date'], format='%Y-%m-%d')
	termdata['ad'] = pd.to_datetime(termdata['Assignment Last Update Date'], format='%Y-%m-%d %H:%M:%S')

	# add column with time difference
	termdata['Gap (# of days)'] = termdata['ad'] - termdata['td']

	# add new org column
	termdata['Org'] = termdata.apply(lambda row: determineOrganization(row['Job Title'], row['Department Name'], row['Location Name']), axis=1)

	# add is leader
	termdata['IsLeader?'] = termdata.apply(lambda row: determineIsLeader(row['Job Title']), axis=1)

	# add is leader or home office
	termdata['IsLeaderOrHomeOffice?'] = (termdata['IsLeader?'] == True) | (termdata['Org'] == 'Home Office')

	# filter out rows by dates
	startdate = pd.Timestamp(date.fromisoformat(stdate))
	termdata_filtered = termdata[(termdata['td'] >= startdate) | (termdata['ad'] >= startdate)]
	enddate = pd.Timestamp(date.fromisoformat(eddate))
    	termdata_filtered = termdata_filtered[termdata_filtered['td'] <= enddate]    

	# filter for rows > 5 days
	termdata_filtered = termdata_filtered[termdata_filtered['Gap (# of days)'] > pd.Timedelta(timedelta(days=5))]

	## process BPM file
	# only look at approved transactions - for now ignore this criteria
	# bpmdata_filtered = bpmdata[bpmdata['Task Outcome'] == 'Approve']
	bpmdata_filtered = bpmdata

	# extract columns
	subjects = bpmdata_filtered['Subject'].tolist()
	dates = bpmdata_filtered['Completion Date'].tolist()

	# create mapping from assoc_id to first completion date and name
	mapping = dict()

	for ind in range(len(subjects)):
	    name, aid = extractTerminationNameAndId(subjects[ind])
	    if aid is None:
	        continue
	    
	    parsed_date = datetime.strptime(dates[ind], '%m/%d/%y')
	    
	    if aid in mapping:
	        # check if existing date is before or after current one
	        existing = mapping[aid]
	        if existing[1] <= parsed_date:
	            continue
	    
	    mapping[aid] = (name, parsed_date)

	# looks up each associate in the levisay report, and finds their bpm term date
	missing = 0
	bpmdatecol = []
	for index, row in termdata_filtered.iterrows():
	    pid = row['Person Number']
	    
	    val = mapping.get(str(pid))
	    if val is None:
	        print(pid, row['First Name'], row['Last Name'])
	        missing += 1
		bpmdatecol.append(datetime(1900,1,1))
		continue
	    bpmdatecol.append(val[1])
	
	
	termdata_filtered.insert(len(termdata_filtered.columns), 'bpmd', bpmdatecol)   
	termdata_filtered['BPM Lookup Date'] = termdata_filtered.apply(lambda x: datetime.strftime(x['bpmd'], "%Y-%m-%d"), axis=1)

	# create new column comparing BPM date to Termination date
	termdata_filtered['Gap (against BPM)'] = termdata_filtered['bpmd'] - termdata_filtered['td']
	termdata_filtered['Possible false positive'] = (termdata_filtered['Gap (against BPM)'] <= pd.Timedelta(timedelta(days=5))) & (termdata_filtered['Gap (against BPM)'] > pd.Timedelta(timedelta(days=0)))

	## process the last logon file
	# extract columns
	userid = logondata['User Name'].tolist()
	logontime = logondata['Logon Time'].tolist()

	mapping = dict()

	for ind in range(len(userid)):
	    username = userid[ind]
	    parsed_date = datetime.strptime(logontime[ind], '%b %d,%Y %I:%M:%S %p')
	    
	    mapping[username] = parsed_date

	# looks up each associate in levisay report, and finds their last logon date
	missing = 0
	logondatecol = []
	for index, row in termdata_filtered.iterrows():
	    username = row['AD Account Name']
	    
	    val = mapping.get(str(username))
	    logondatecol.append(val)
	    if val is None:
	        missing += 1

	termdata_filtered.insert(len(termdata_filtered.columns), 'logond', logondatecol)
	termdata_filtered['Last Login Date'] = termdata_filtered.apply(lambda x: datetime.strftime(x['logond'], "%Y-%m-%d") if not pd.isnull(x['logond']) else "N/A", axis=1)
	termdata_filtered['Last Login Compared to Term Date'] = termdata_filtered['logond'] - termdata_filtered['td']

	morethan5 = "Logged in More than 5 Days after term date"
	within5 = "Logged in after term date but, within 5 days"
	notwithin = "No"

	termdata_filtered['Late Login?'] = termdata_filtered.apply(lambda x: morethan5 if x['Last Login Compared to Term Date'] > pd.Timedelta(timedelta(days=5)) else (within5 if x['Last Login Compared to Term Date'] > pd.Timedelta(timedelta(days=0)) else notwithin), axis=1)
	termdata_filtered['Last Login Compared to Term Date'] = termdata_filtered.apply(lambda x: str(x['Last Login Compared to Term Date']) if not pd.isnull(x['Last Login Compared to Term Date']) else "N/A" , axis=1)

	## process the all users file
	userid = usersdata['SAM Account Name'].tolist()
	status = usersdata['Account Status'].tolist()

	accstatus_mapping = dict(zip(userid, status))
	    
	# looks up each associate and finds status
	missing = 0
	accstatuscol = []
	for index, row in termdata_filtered.iterrows():
	    username = row['AD Account Name']
	    val = accstatus_mapping.get(str(username))
	    accstatuscol.append(val)
	    if val is None:
	        missing += 1
	        
	termdata_filtered.insert(len(termdata_filtered.columns), 'Account Status', accstatuscol)

	## process the recently disabled file
	# extract columns
	userid = disableddata['User Name'].tolist()
	disabledtime = disableddata['Modified Time'].tolist()

	mapping = dict()

	for ind in range(len(userid)):
	    username = userid[ind]
	    parsed_date = datetime.strptime(disabledtime[ind], '%b %d,%Y %I:%M:%S %p')
	    
	    if username in mapping:
	        # check if existing date is before or after current one
	        existing = mapping[username]
	        if existing > parsed_date:
	            continue
	    
	    mapping[username] = parsed_date

	# looks up each associate in levisay report, and finds their latest disable date
	missing = 0
	disableddatecol = []
	for index, row in termdata_filtered.iterrows():
	    username = row['AD Account Name']
	    
	    val = mapping.get(str(username))
	    disableddatecol.append(val)
	    if val is None:
	        missing += 1
	        
	termdata_filtered.insert(len(termdata_filtered.columns), 'dd', disableddatecol)
	termdata_filtered['Disable Date'] = termdata_filtered.apply(lambda x: datetime.strftime(x['dd'], "%Y-%m-%d") if not pd.isnull(x['dd']) else "N/A", axis=1)

	# gap between disabled date and last update date
	termdata_filtered['Access Disabled Date compared to Last Update Date'] = termdata_filtered['dd'] - termdata_filtered['ad']
	termdata_filtered['Access Disabled Date compared to Last Update Date'] = termdata_filtered.apply(lambda x: str(x['Access Disabled Date compared to Last Update Date']) if not pd.isnull(x['Access Disabled Date compared to Last Update Date']) else "N/A", axis=1)

	# add final group column
	groupstrings = ['1. No Activity After Term Date', '2. Activity After Term Date']
	termdata_filtered['GROUP'] = termdata_filtered.apply(lambda x: groupstrings[0] if x['Late Login?'] == notwithin else groupstrings[1], axis=1)

	# remove temporary columns
	termdata_filtered = termdata_filtered.drop(['td','ad','bpmd', 'logond', 'dd'], axis=1)

	# computes stats based on Org
	orgs = termdata_filtered['Org'].tolist()
	counts = dict()
	total = 0
	for a in orgs:
	    if a in counts:
	        counts[a] += 1
	    else:
	        counts[a] = 1
	    total += 1

	print(counts)
	print("Total: %d" % total)

	leaders = termdata_filtered['IsLeaderOrHomeOffice?'].tolist()
	counts_leader = dict()
	total_leaders = 0
	for ind in range(len(leaders)):
	    isleader = leaders[ind]
	    org = orgs[ind]
	    
	    if isleader:
	        if org in counts_leader:
	            counts_leader[org] += 1
	        else:
	            counts_leader[org] = 1    
	        total_leaders += 1

	print(counts_leader)
	print("Total leaders: %d" % total_leaders)

	# save to disk
	termdata_filtered.to_excel('output_ELC27.xlsx')

