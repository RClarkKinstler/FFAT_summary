'''******************************************************************************
 Copyright 2014-2019 IDEX ASA. All Rights Reserved. www.idexbiometrics.com

 IDEX ASA is the owner of this software and all intellectual property rights
 in and to the software. The software may only be used together with IDEX
 fingerprint sensors, unless otherwise permitted by IDEX ASA in writing.

 This copyright notice must not be altered or removed from the software.

 DISCLAIMER OF WARRANTY/LIMITATION OF REMEDIES: unless otherwise agreed, IDEX
 ASA has no obligation to support this software, and the software is provided
 "AS IS", with no express or implied warranties of any kind, and IDEX ASA is
 not to be liable for any damages, any relief, or for any claim by any third
 party, arising from use of this software.

 Image capture and processing logic is defined and controlled by IDEX ASA in
 order to maximize FAR / FRR performance.
******************************************************************************'''
import os
import os.path
import csv
import argparse
import sys
import re
import FFAT_watermark_report

def summary_report( csvName, logName, add_1=False, addFTA2FRR = False, xlsx=False, previous=None ):
    if xlsx :
        try:
            xw = __import__( 'xlsxwriter', globals(), locals())
        except:
            print('Import of xlsxwriter failed.  Maybe it is not installed?')
            return 1
    if add_1 :
        tblIncr = 1
    else :
        tblIncr = 0
    inFile = open( csvName, newline='')
    reader = csv.reader( inFile)

    #  Search for matcher and subtype in Command Line
    cmdParser = argparse.ArgumentParser( )
    cmdParser.add_argument( 'ffatExe' )
    cmdParser.add_argument( 'db' )
    cmdParser.add_argument( '--matcherID')
    cmdParser.add_argument( '--matcherSubtype')
    matcherFound = False
    avgFAMatch = '-1'
    for row in  reader:
        if row[0] == 'Tool Version' :
            version = row[1].split()[-1]
        elif row[0] == 'Command Line' :
            settings = row[1]
            knowns,unknowns = cmdParser.parse_known_args( row[1].split())
            matcher = knowns.matcherID
            subtype = knowns.matcherSubtype
            if matcher :
                matcherFound = True
            if subtype == None :
                subtype = '0'
        elif row[0].startswith('Average per Enroll') :
            avgEnroll = row[1]
        elif row[0].startswith('Maximum per Enroll') :
            maxEnroll = row[1]
        elif row[0].startswith('Minimum per Enroll') :
            minEnroll = row[1]
        elif row[0].startswith('Average per Match') :
            avgMatch = row[1]
        elif row[0].startswith('Maximum per Match') :
            maxMatch = row[1]
        elif row[0].startswith('Minimum per Match') :
            minMatch = row[1]
        elif row[0].startswith('Average per FA Match') :
            avgFAMatch = row[1]
        elif row[0].startswith('Maximum per FA Match') :
            maxFAMatch = row[1]
        elif row[0].startswith('Minimum per FA Match') :
            minFAMatch = row[1]
        elif row[0].startswith('Average per FR Match') :
            avgFRMatch = row[1]
        elif row[0].startswith('Maximum per FR Match') :
            maxFRMatch = row[1]
        elif row[0].startswith('Minimum per FR Match') :
            minFRMatch = row[1]
        elif row[0].startswith('Output Directory') :
            outputDir = row[1]
        elif row[0].startswith('Input Directory') :
            db = row[1]
        elif row[0].startswith('Elapsed time') :
            elapsedTime = row[1]
            break
    if not matcherFound :
        for row in reader :
            if len(row) > 1 and row[1].startswith('MatcherID') :
                candidate = next(reader)
                if candidate[1] == 'Events' :
                    continue
                else :
                    matcher = candidate[1]
                    subtype = candidate[2]
                    matcherFound = True
                    break
    if not matcherFound :
        print( 'failed to find matcher.  Exiting')
        return 1
    ftarValue = 0.0
    for row in reader :
        if row[0].startswith("Match Threshold") :
            threshold = row[1]
        elif row[0].startswith( 'Total number of Users') :
            totalUsers = row[1]
        elif row[0].startswith( 'Total number of Fingers') :
            totalFingers = row[1]
        elif row[0].startswith( 'Total number of Enroll images') :
            totalEnrollImages = row[1]
        elif row[0].startswith( 'Total number of images for FA') :
            totalFAimages = row[1]
        elif row[0].startswith( 'Total number of images for FR') :
            totalFRimages = row[1]
        elif row[0].startswith( 'Finger FTER') :
            if len( row) < 6 :
                continue
            fingerFTER = row[1] + ' % (' + row[3] + ' failures in ' + row[5] + ' attempts)'
        elif row[0].startswith( 'User FTER') :
            if len( row) < 6 :
                continue
            userFTER = row[1] + ' % (' + row[3] + ' failures in ' + row[5] + ' attempts)'
        elif row[0].startswith( 'TemplateCreator FTAR') :
            templateFTAR = row[1] + ' % (' + row[3] + ' failures in ' + row[5] + ' attempts)'
        elif row[0].startswith( 'CrossMatcher FTAR') :
            crossmatcherFTAR = row[1] + ' % (' + row[3] + ' failures in ' + row[5] + ' attempts)'
            if addFTA2FRR :
                ftarValue = float( row[1])
        elif row[0].startswith( 'Total FTAR') :
            totalFTAR = row[1] + ' % (' + row[3] + ' failures in ' + row[5] + ' attempts)'

        elif row[0].startswith("False Accepts") :
            fas = row[1]
        elif row[0].startswith("FAR (False Accept Rate)") :
            far = row[1]
        elif row[0].startswith("False Rejects") :
            frs = row[1]
        elif row[0].startswith("FRR (False Reject Rate)") :
            frr = row[1]
        elif row[0].startswith("Transactional False Rejects") :
            tfrs = row[1]
        elif row[0].startswith("Transactional FRR") :
            tfrr = row[1]
            break
    if float(far) == 0.0:
        farnnkStr = '(1 : INF)'
    else:
        farnnk = round( 0.1 / float(far), 1)
        farnnkStr = '(1 : ' + str(farnnk) + ' K)'
    #  Get the 1:nnK FAR data
    if previous :
        inResultTable = False
        farTbl = []
        for row in reader :
            if row and row[0].startswith('Crossmatcher Results') :
                next(reader)
                inResultTable = True
                break
        if not inResultTable :
            print( 'Crossmatcher Results table not found.  Quitting.')
            return 1
        for row in reader :
            if len( row) == 0 :
                break
            else :
                farTbl.append( row)
        outRows = {}
        outNums = {}
        labels =  [ 'Calculated FAR Level', 'Score', 'FA Count', 'FAR', 'FR Count', 'FRR']

        for prevEntry in previous :
            for line in farTbl :
                if prevEntry[1] == 'INF' or line[4] == 'INF' :
                    outRows[prevEntry[0]] = ['FRR at Calculated 1:' + prevEntry[0] +' FAR']
                    outNums[prevEntry[0]] = ['FRR at Calculated 1:' + prevEntry[0] +' FAR']
                    break
                if int(line[0])+tblIncr >= int( prevEntry[1]) :
                    # FRR at Calculated 1:1K FAR, 1, 88, 0.00648% (1 : 15.4 K), 1186, 13.97102%
                    outRows[prevEntry[0]] = ['FRR at Calculated 1:'+ prevEntry[0]+' FAR', str(int(line[0])+tblIncr), line[1], line[2]+'% ('+
                    line[4]+')', line[5], str( float( line[6]) + ftarValue)+'%']
                    outNums[prevEntry[0]] = ['FRR at Calculated 1:'+prevEntry[0]+' FAR', int(line[0])+tblIncr, int( line[1]), line[2]+'% ('+
                    line[4]+')', int( line[5]), str( float( line[6]) + ftarValue)+'%']
                    break
    else :
        inResultTable = False
        farTbl = []
        for row in reader :
            if row and row[0].startswith('Crossmatcher Results') :
                next(reader)
                inResultTable = True
                break
        if not inResultTable :
            inFile.close()
            inFile = open( csvName, newline='')
            reader = csv.reader( inFile)
            for row in reader :
                if row[0].startswith('1:nnK FAR Transitions Tables') :
                    next(reader)
                    inResultTable = True
                    break
        if inResultTable :
            for row in reader :
                if len( row) == 0 :
                    break
                else :
                    farTbl.append( row)
            line1k = line5k = line10k = line25k = line30k = line40k = line50k = line75k = line100k = line150k = line200k = line250k = line300k = line400k = line500k = line750k = line1m = line2m = line3m = line5m = None
            for line in farTbl :
                if line[4] == 'INF' or (float( line[4].split()[2]) >= 1.0 and line[4].split()[-1] == 'K') :
                    line1k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 5.0 and line[4].split()[-1] == 'K' :
                    line5k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 10.0 and line[4].split()[-1] == 'K' :
                    line10k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 25.0 and line[4].split()[-1] == 'K' :
                    line25k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 30.0 and line[4].split()[-1] == 'K' :
                    line30k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 40.0 and line[4].split()[-1] == 'K' :
                    line40k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 50.0 and line[4].split()[-1] == 'K' :
                    line50k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 75.0 and line[4].split()[-1] == 'K' :
                    line75k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 100.0 and line[4].split()[-1] == 'K' :
                    line100k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 150.0 and line[4].split()[-1] == 'K' :
                    line150k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 200.0 and line[4].split()[-1] == 'K' :
                    line200k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 250.0 and line[4].split()[-1] == 'K' :
                    line250k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 300.0 and line[4].split()[-1] == 'K' :
                    line300k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 400.0 and line[4].split()[-1] == 'K' :
                    line400k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 500.0 and line[4].split()[-1] == 'K' :
                    line500k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 750.0 and line[4].split()[-1] == 'K' :
                    line750k = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 1000.0 and line[4].split()[-1] == 'K' :
                    line1m = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 2000.0 and line[4].split()[-1] == 'K' :
                    line2m = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 3000.0 and line[4].split()[-1] == 'K' :
                    line3m = line
                    break
            for line in farTbl :
                if line[4] == 'INF' :
                    break
                if float( line[4].split()[2]) >= 5000.0 and line[4].split()[-1] == 'K' :
                    line5m = line
                    break
            if line1k or line5k or line10k or line25k or line50k or line100k or line150k or line200k or line250k or line300k or line400k or line500k or line750k or line1m or line2m or line3m or line5m:
                labels =  [ 'Calculated FAR Level', 'Score', 'FA Count', 'FAR', 'FR Count', 'FRR']
            #
            #  Conditionally add 1 to all of the score values because of a bug in FFAT
            #
            if line1k :
                row1k = ['FRR at Calculated 1:1K FAR', str( int(line1k[0])+tblIncr), line1k[1], line1k[2]+'% ('+line1k[4]+')', line1k[5], str( float( line1k[6]) + ftarValue)+'%']
                num1k = ['FRR at Calculated 1:1K FAR', int( line1k[0])+tblIncr, int( line1k[1]), line1k[2]+'% ('+line1k[4]+')', int( line1k[5]), str( float( line1k[6]) + ftarValue)+'%']
            else :
                row1k = num1k = ['FRR at Calculated 1:1K FAR']

            if line5k :
                row5k = ['FRR at Calculated 1:5K FAR', str( int(line5k[0])+tblIncr), line5k[1], line5k[2]+'% ('+line5k[4]+')', line5k[5], str( float( line5k[6]) + ftarValue)+'%']
                num5k = ['FRR at Calculated 1:5K FAR', int( line5k[0])+tblIncr, int( line5k[1]), line5k[2]+'% ('+line5k[4]+')', int( line5k[5]), str( float( line5k[6]) + ftarValue)+'%']
            else :
                row5k = num5k = ['FRR at Calculated 1:5K FAR']

            if line10k :
                row10k = ['FRR at Calculated 1:10K FAR', str( int(line10k[0])+tblIncr), line10k[1], line10k[2]+'% ('+line10k[4]+')', line10k[5], str( float( line10k[6]) + ftarValue)+'%']
                num10k = ['FRR at Calculated 1:10K FAR', int( line10k[0])+tblIncr, int( line10k[1]), line10k[2]+'% ('+line10k[4]+')', int( line10k[5]), str( float( line10k[6]) + ftarValue)+'%']
            else :
                row10k = num10k = ['FRR at Calculated 1:10K FAR']

            if line25k :
                row25k = ['FRR at Calculated 1:25K FAR', str( int(line25k[0])+tblIncr), line25k[1], line25k[2]+'% ('+line25k[4]+')', line25k[5], str( float( line25k[6]) + ftarValue)+'%']
                num25k = ['FRR at Calculated 1:25K FAR', int( line25k[0])+tblIncr, int( line25k[1]), line25k[2]+'% ('+line25k[4]+')', int( line25k[5]), str( float( line25k[6]) + ftarValue)+'%']
            else :
                row25k = num25k = ['FRR at Calculated 1:25K FAR']

            if line30k :
                row30k = ['FRR at Calculated 1:30K FAR', str( int(line30k[0])+tblIncr), line30k[1], line30k[2]+'% ('+line30k[4]+')', line30k[5], str( float( line30k[6]) + ftarValue)+'%']
                num30k = ['FRR at Calculated 1:30K FAR', int( line30k[0])+tblIncr, int( line30k[1]), line30k[2]+'% ('+line30k[4]+')', int( line30k[5]), str( float( line30k[6]) + ftarValue)+'%']
            else :
                row30k = num30k = ['FRR at Calculated 1:30K FAR']

            if line40k :
                row40k = ['FRR at Calculated 1:40K FAR', str( int(line40k[0])+tblIncr), line40k[1], line40k[2]+'% ('+line40k[4]+')', line40k[5], str( float( line40k[6]) + ftarValue)+'%']
                num40k = ['FRR at Calculated 1:40K FAR', int( line40k[0])+tblIncr, int( line40k[1]), line40k[2]+'% ('+line40k[4]+')', int( line40k[5]), str( float( line40k[6]) + ftarValue)+'%']
            else :
                row40k = num40k = ['FRR at Calculated 1:40K FAR']

            if line50k :
                row50k = ['FRR at Calculated 1:50K FAR', str( int(line50k[0])+tblIncr), line50k[1], line50k[2]+'% ('+line50k[4]+')', line50k[5], str( float( line50k[6]) + ftarValue)+'%']
                num50k = ['FRR at Calculated 1:50K FAR', int( line50k[0])+tblIncr, int( line50k[1]), line50k[2]+'% ('+line50k[4]+')', int( line50k[5]), str( float( line50k[6]) + ftarValue)+'%']
            else :
                row50k = num50k = ['FRR at Calculated 1:50K FAR']

            if line75k :
                row75k = ['FRR at Calculated 1:75K FAR', str( int(line75k[0])+tblIncr), line75k[1], line75k[2]+'% ('+line75k[4]+')', line75k[5], str( float( line75k[6]) + ftarValue)+'%']
                num75k = ['FRR at Calculated 1:75K FAR', int( line75k[0])+tblIncr, int( line75k[1]), line75k[2]+'% ('+line75k[4]+')', int( line75k[5]), str( float( line75k[6]) + ftarValue)+'%']
            else :
                row75k = num75k = ['FRR at Calculated 1:75K FAR']

            if line100k :
                row100k = ['FRR at Calculated 1:100K FAR', str( int(line100k[0])+tblIncr), line100k[1], line100k[2]+'% ('+line100k[4]+')', line100k[5], str( float( line100k[6]) + ftarValue)+'%']
                num100k = ['FRR at Calculated 1:100K FAR', int( line100k[0])+tblIncr, int( line100k[1]), line100k[2]+'% ('+line100k[4]+')', int( line100k[5]), str( float( line100k[6]) + ftarValue)+'%']
            else :
                row100k = num100k = ['FRR at Calculated 1:100K FAR']

            if line150k :
                row150k = ['FRR at Calculated 1:150K FAR', str( int(line150k[0])+tblIncr), line150k[1], line150k[2]+'% ('+line150k[4]+')', line150k[5], str( float( line150k[6]) + ftarValue)+'%']
                num150k = ['FRR at Calculated 1:150K FAR', int( line150k[0])+tblIncr, int( line150k[1]), line150k[2]+'% ('+line150k[4]+')', int( line150k[5]), str( float( line150k[6]) + ftarValue)+'%']
            else :
                row150k = num150k = ['FRR at Calculated 1:150K FAR']

            if line200k :
                row200k = ['FRR at Calculated 1:200K FAR', str( int(line200k[0])+tblIncr), line200k[1], line200k[2]+'% ('+line200k[4]+')', line200k[5], str( float( line200k[6]) + ftarValue)+'%']
                num200k = ['FRR at Calculated 1:200K FAR', int( line200k[0])+tblIncr, int( line200k[1]), line200k[2]+'% ('+line200k[4]+')', int( line200k[5]), str( float( line200k[6]) + ftarValue)+'%']
            else :
                row200k = num200k = ['FRR at Calculated 1:200K FAR']

            if line250k :
                row250k = ['FRR at Calculated 1:250K FAR', str( int(line250k[0])+tblIncr), line250k[1], line250k[2]+'% ('+line250k[4]+')', line250k[5], str( float( line250k[6]) + ftarValue)+'%']
                num250k = ['FRR at Calculated 1:250K FAR', int( line250k[0])+tblIncr, int( line250k[1]), line250k[2]+'% ('+line250k[4]+')', int( line250k[5]), str( float( line250k[6]) + ftarValue)+'%']
            else :
                row250k = num250k = ['FRR at Calculated 1:250K FAR']

            if line300k :
                row300k = ['FRR at Calculated 1:300K FAR', str( int(line300k[0])+tblIncr), line300k[1], line300k[2]+'% ('+line300k[4]+')', line300k[5], str( float( line300k[6]) + ftarValue)+'%']
                num300k = ['FRR at Calculated 1:300K FAR', int( line300k[0])+tblIncr, int( line300k[1]), line300k[2]+'% ('+line300k[4]+')', int( line300k[5]), str( float( line300k[6]) + ftarValue)+'%']
            else :
                row300k = num300k = ['FRR at Calculated 1:300K FAR']

            if line400k :
                row400k = ['FRR at Calculated 1:400K FAR', str( int(line400k[0])+tblIncr), line400k[1], line400k[2]+'% ('+line400k[4]+')', line400k[5], str( float( line400k[6]) + ftarValue)+'%']
                num400k = ['FRR at Calculated 1:400K FAR', int( line400k[0])+tblIncr, int( line400k[1]), line400k[2]+'% ('+line400k[4]+')', int( line400k[5]), str( float( line400k[6]) + ftarValue)+'%']
            else :
                row400k = num400k = ['FRR at Calculated 1:400K FAR']

            if line500k :
                row500k = ['FRR at Calculated 1:500K FAR', str( int(line500k[0])+tblIncr), line500k[1], line500k[2]+'% ('+line500k[4]+')', line500k[5], str( float( line500k[6]) + ftarValue)+'%']
                num500k = ['FRR at Calculated 1:500K FAR', int( line500k[0])+tblIncr, int( line500k[1]), line500k[2]+'% ('+line500k[4]+')', int( line500k[5]), str( float( line500k[6]) + ftarValue)+'%']
            else :
                row500k = num500k = ['FRR at Calculated 1:500K FAR']

            if line750k :
                row750k = ['FRR at Calculated 1:750K FAR', str( int(line750k[0])+tblIncr), line750k[1], line750k[2]+'% ('+line750k[4]+')', line750k[5], str( float( line750k[6]) + ftarValue)+'%']
                num750k = ['FRR at Calculated 1:750K FAR', int( line750k[0])+tblIncr, int( line750k[1]), line750k[2]+'% ('+line750k[4]+')', int( line750k[5]), str( float( line750k[6]) + ftarValue)+'%']
            else :
                row750k = num750k = ['FRR at Calculated 1:750K FAR']

            if line1m :
                row1m = ['FRR at Calculated 1:1M FAR', str( int(line1m[0])+tblIncr), line1m[1], line1m[2]+'% ('+line1m[4]+')', line1m[5], str( float( line1m[6]) + ftarValue)+'%']
                num1m = ['FRR at Calculated 1:1M FAR', int( line1m[0])+tblIncr, int( line1m[1]), line1m[2]+'% ('+line1m[4]+')', int( line1m[5]), str( float( line1m[6]) + ftarValue)+'%']
            else :
                row1m = num1m = ['FRR at Calculated 1:1M FAR']

            if line2m :
                row2m = ['FRR at Calculated 1:2M FAR', str( int(line2m[0])+tblIncr), line2m[1], line2m[2]+'% ('+line2m[4]+')', line2m[5], str( float( line2m[6]) + ftarValue)+'%']
                num2m = ['FRR at Calculated 1:2M FAR', int( line2m[0])+tblIncr, int( line2m[1]), line2m[2]+'% ('+line2m[4]+')', int( line2m[5]), str( float( line2m[6]) + ftarValue)+'%']
            else :
                row2m = num2m = ['FRR at Calculated 1:2M FAR']

            if line3m :
                row3m = ['FRR at Calculated 1:3M FAR', str( int(line3m[0])+tblIncr), line3m[1], line3m[2]+'% ('+line3m[4]+')', line3m[5], str( float( line3m[6]) + ftarValue)+'%']
                num3m = ['FRR at Calculated 1:3M FAR', int( line3m[0])+tblIncr, int( line3m[1]), line3m[2]+'% ('+line3m[4]+')', int( line3m[5]), str( float( line3m[6]) + ftarValue)+'%']
            else :
                row3m = num3m = ['FRR at Calculated 1:3M FAR']

            if line5m :
                row5m = ['FRR at Calculated 1:5M FAR', str( int(line5m[0])+tblIncr), line5m[1], line5m[2]+'% ('+line5m[4]+')', line5m[5], str( float( line5m[6]) + ftarValue)+'%']
                num5m = ['FRR at Calculated 1:5M FAR', int( line5m[0])+tblIncr, int( line5m[1]), line5m[2]+'% ('+line5m[4]+')', int( line5m[5]), str( float( line5m[6]) + ftarValue)+'%']
            else :
                row5m = num5m = ['FRR at Calculated 1:5M FAR']
        inFile.close()
        inFile = open( csvName, newline='')
        reader = csv.reader( inFile)
        inHistogram = False
        for row in reader :
            if len(row) == 0 :
                continue
            if row[0].startswith('Histogram of Finger Sample Matched Index') :
                inHistogram = True
                break
        if inHistogram :
            for row in reader :
                if len( row) == 0 :
                    break
                elif row[0].startswith('Total Found') :
                    histTotal = row[1]
                elif row[0].startswith('Min Found') :
                    histMin = row[2].split()[-1]
                elif row[0].startswith('Max Found') :
                    histMax = row[2].split()[-1]
                elif row[0].startswith('Max Possible') :
                    histPossible = row[2].split()[-1]
                elif row[0].startswith('Average') :
                    histAvg = row[1]
                elif row[0].startswith('StdDev') :
                    histStdDev = row[1]
                elif row[0].startswith('Finger') :
                    histTitles = row
                    histTable = []
                    while True :
                        row = next(reader)
                        if row[0].startswith('-----') :
                            break
                        histTable.append( ['0', row[0], row[1], row[3][:-1], row[4][:-1], row[5][:-1]])
                else :
                    break
        else :
            print( 'Histogram of Finger Sample Matched Index not found.  Quitting')
            return 1
        inFile.close()

    argsPat = r'\| All FFAT Arguments \|'
    updPat = r'updateTemplates = (\w+)'
    maxUpdPat = r'maxTemplateUpdates = (-?\d+)'
    maxEnrollImgPat = r'maxEnrollImages = (-?\d+)'
    argsValPat = r'\| Argument Validation \|'

    inArguments = False
    updTempl = 'unknown'
    maxUpdates = ''
    for line in open( logName) :
        if not inArguments :
            argsMatch = re.search( argsPat, line)
            if argsMatch :
                inArguments = True
        else :
            updMatch = re.search( updPat, line)
            if updMatch :
                updTempl = updMatch.group( 1)
            else :
                maxUpdMatch = re.search( maxUpdPat, line)
                if maxUpdMatch :
                    maxUpdates = maxUpdMatch.group( 1)
                else :
                    maxEnrollImgMatch = re.search( maxEnrollImgPat, line)
                    if maxEnrollImgMatch :
                        maxEnrollImages = maxEnrollImgMatch.group( 1)
                    elif re.search( argsValPat, line) :
                        break
    if updTempl == 'false' :
        updTempl = 'disabled'
        maxUpdates = 'disabled'
    else :
        updTempl = 'enabled'
        if maxUpdates == '-1' :
            maxUpdates = 'maximum'
    # Get max enrollment images.  Call it enrollImageMax
    procDirPat = r'Processing directory (.*)'
    for line in open( logName) :
        procDirMatch = re.search( procDirPat, line)
        if procDirMatch :
            procDir = procDirMatch.group( 1)
            break
    if os.path.exists( procDir) :
        users = [x for x in os.listdir( procDir) if os.path.isdir( os.path.join( procDir, x))]
        userDir = os.path.join( procDir, users[0])
        if os.path.exists( userDir) :
            fingers = [x for x in os.listdir( userDir) if os.path.isdir( os.path.join( userDir, x))]
            enrollDir = os.path.join( userDir, fingers[0], 'enroll')
            images = [x for x in os.listdir( enrollDir) if re.search( r'(?i)\.BMP$', x)]
            imgCount = len( images)
            if maxEnrollImages == '-1' :
                enrollImageMax = imgCount
            else :
                enrollImageMax = min( imgCount, int( maxEnrollImages))
        else :
            enrollImageMax = maxEnrollImages
    else :
        enrollImageMax = maxEnrollImages
    #  See if we have optical (enroll) image files.
    csvDir = os.path.dirname( csvName)
    subdirs = [x for x in os.listdir( csvDir) if os.path.isdir( os.path.join( csvDir, x))]
    userDir = os.path.join( csvDir, subdirs[0])
    subdirs = [x for x in os.listdir( userDir) if os.path.isdir( os.path.join( userDir, x))]
    enrollDir = os.path.join( userDir, subdirs[0], 'enroll')
    if not os.path.isdir( enrollDir) :
        print( 'enroll folder not found in results.')
        return 1
    images = [x for x in os.listdir( enrollDir) if re.search( r'(?i)\.BMP$', x)]
    opticalFound = False
    for imgFile in images :
        stats = os.stat( os.path.join( enrollDir, imgFile))
        if stats.st_size > 25000 :
            opticalFound = True
    if opticalFound :
        maxPrimaryImages = '0'
        startSecondary = '0'
        maxPrimaryPat = r'maxPrimaryImages = (.*)$'
        startSecondaryPat = r'startSecondaryAtNextImage = (.*)$'
        for line in open( logName) :
            maxPrimaryMatch = re.search( maxPrimaryPat, line)
            if maxPrimaryMatch :
                maxPrimaryImages = maxPrimaryMatch.group( 1)
            startSecondaryMatch = re.search( startSecondaryPat, line)
            if startSecondaryMatch :
                startSecondary = startSecondaryMatch.group( 1)
                break

    if xlsx :
        workbook = xw.Workbook( os.path.splitext( csvName)[0] +'_summary.xlsx')
        allSheet = workbook.add_worksheet( )
        allSheet.set_column( 0, 0, 40)
        allSheet.set_column( 3, 3, 20)
        allSheet.write( 0, 0, 'DB/Input Directory')
        allSheet.merge_range( 'B1:R1', db)
        allSheet.write( 1, 0, 'Output Directory')
        allSheet.merge_range( 'B2:R2', outputDir)
        allSheet.write( 2, 0, 'FFAT version')
        allSheet.write( 2, 1, version)
        allSheet.write( 3, 0, 'FFAT settings')
        allSheet.merge_range( 'B4:R4', settings)
        allSheet.write( 4, 0, 'Matcher')
        allSheet.write( 4, 1,  ' '.join( [matcher, subtype]))
        allSheet.write( 5, 0, 'Total number of Users')
        allSheet.write( 5, 1, int( totalUsers))
        allSheet.write( 6, 0, 'Total number of Fingers')
        allSheet.write( 6, 1, int( totalFingers))
        allSheet.write( 7, 0, 'Total number of Enroll images')
        allSheet.write( 7, 1, int( totalEnrollImages))
        allSheet.write( 8, 0, 'Total number of images for FA')
        allSheet.write( 8, 1, int( totalFAimages))
        allSheet.write( 9, 0, 'Total number of images for FR')
        allSheet.write( 9, 1, int( totalFRimages))
        allSheet.write( 10, 0, 'Finger FTER')
        allSheet.write( 10, 1, fingerFTER)
        allSheet.write( 11, 0, 'User FTER')
        allSheet.write( 11, 1, userFTER)
        allSheet.write( 12, 0, 'TemplateCreator FTAR')
        allSheet.write( 12, 1, templateFTAR)
        allSheet.write( 13, 0, 'CrossMatcher FTAR')
        allSheet.write( 13, 1, crossmatcherFTAR)
        allSheet.write( 14, 0, 'Total FTAR')
        allSheet.write( 14, 1, totalFTAR)
        rowNum = 15
        if previous :
            for col, label in zip( range( 1, len(labels)), labels[1:]) :
                allSheet.write( rowNum, col, label)
            rowNum += 1
            rowKeys = ['1K', '5K', '10K', '25K', '30K', '40K', '50K', '75K', '100K', '150K', '200K', '250K', '300K', '400K', '500K', '750K', '1M', '2M', '3M', '5M']
            for rowKey in rowKeys :
                for col, value in zip( range( len( outNums[rowKey])), outNums[rowKey]) :
                    allSheet.write( rowNum, col, value)
                rowNum += 1

        elif inResultTable and line1k :
            for col, label in zip( range( 1, len(labels)), labels[1:]) :
                allSheet.write( rowNum, col, label)
            rowNum += 1
            for col, value in zip( range( len( num1k)), num1k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num5k)), num5k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num10k)), num10k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num25k)), num25k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num30k)), num30k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num40k)), num40k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num50k)), num50k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num75k)), num75k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num100k)), num100k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num150k)), num150k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num200k)), num200k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num250k)), num250k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num300k)), num300k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num400k)), num400k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num500k)), num500k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num750k)), num750k) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num1m)), num1m) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num2m)), num2m) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num3m)), num3m) :
                allSheet.write( rowNum, col, value)
            rowNum += 1
            for col, value in zip( range( len( num5m)), num5m) :
                allSheet.write( rowNum, col, value)
            rowNum += 1

        allSheet.write( rowNum, 0, 'FAR/FRR at Executed Threshold')
        allSheet.write( rowNum, 1, int( threshold))
        allSheet.write( rowNum, 2, int( fas))
        allSheet.write( rowNum, 3,  far+'%'+farnnkStr)
        allSheet.write( rowNum, 4, int( frs))
        allSheet.write( rowNum, 5, frr+'%')
        rowNum += 1
        allSheet.write( rowNum, 0, 'Transactional FRR at Executed Threshold')
        allSheet.write( rowNum, 1, int( threshold))
        allSheet.write( rowNum, 4, int( tfrs))
        allSheet.write( rowNum, 5, tfrr+'%')
        rowNum += 1
        allSheet.write( rowNum, 0, 'Elapsed Time')
        allSheet.write( rowNum, 1, float( elapsedTime))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Average Enroll Time')
        allSheet.write( rowNum, 1, float( avgEnroll))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Max Enroll Time')
        allSheet.write( rowNum, 1, int( maxEnroll))
        allSheet.write( rowNum, 0, 'Min Enroll Time')
        allSheet.write( rowNum, 1, int( minEnroll))
        rowNum += 1
        if avgFAMatch == '-1' :
            allSheet.write( rowNum, 0, 'Average Match Time')
            allSheet.write( rowNum, 1, float( avgMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Max Match Time')
            allSheet.write( rowNum, 1, int( maxMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Min Match Time')
            allSheet.write( rowNum, 1, int( minMatch))
            rowNum += 1
        else :
            allSheet.write( rowNum, 0, 'Average FA Match Time')
            allSheet.write( rowNum, 1, float( avgFAMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Max FA Match Time')
            allSheet.write( rowNum, 1, int( maxFAMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Min FA Match Time')
            allSheet.write( rowNum, 1, int( minFAMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Average FR Match Time')
            allSheet.write( rowNum, 1, float( avgFRMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Max FR Match Time')
            allSheet.write( rowNum, 1, int( maxFRMatch))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Min FR Match Time')
            allSheet.write( rowNum, 1, int( minFRMatch))
            rowNum += 1
        if enrollImageMax > 0 :
            allSheet.write( rowNum, 0, 'Max Enrollment Images')
            allSheet.write( rowNum, 1, enrollImageMax)
            rowNum += 1
        allSheet.write( rowNum, 0, 'Histogram of Finger Sample Matched Index')
        rowNum += 1
        allSheet.write( rowNum, 0, 'Total Found')
        allSheet.write( rowNum, 1, int(histTotal))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Min Found')
        allSheet.write( rowNum, 1, int(histMin))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Max Found')
        allSheet.write( rowNum, 1, int(histMax))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Max Possible')
        allSheet.write( rowNum, 1, int(histPossible))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Average')
        allSheet.write( rowNum, 1, float(histAvg))
        rowNum += 1
        allSheet.write( rowNum, 0, 'Std Dev')
        allSheet.write( rowNum, 1, float(histStdDev))
        rowNum += 1
        for col, value in zip( range( len( histTitles)), histTitles) :
            allSheet.write( rowNum, col, value)
        rowNum += 1
        for line in histTable :
            for col, valStr in zip( range( len( line)), line) :
                value = int(valStr) if col < 3 else float(valStr)
                allSheet.write( rowNum, col, value)
            rowNum += 1
        allSheet.write( rowNum, 0, 'Update Templates')
        allSheet.write( rowNum, 1, updTempl)
        rowNum += 1
        allSheet.write( rowNum, 0, 'Max Template Updates')
        allSheet.write( rowNum, 1, maxUpdates)
        if opticalFound :
            rowNum += 1
            allSheet.write( rowNum, 0, 'Max Primary Images')
            allSheet.write( rowNum, 1, int( maxPrimaryImages))
            rowNum += 1
            allSheet.write( rowNum, 0, 'Start Secondary At Next Image')
            allSheet.write( rowNum, 1, int( startSecondary))
    else :
        writer = csv.writer( open( os.path.splitext( csvName)[0] + '_summary.csv', 'w', newline=''))
        writer.writerow( ['DB/Input Directory', db])
        writer.writerow( ['Output directory', outputDir])
        writer.writerow( ['FFAT version', version])
        writer.writerow( ['FFAT settings', settings])
        writer.writerow( ['Matcher', ' '.join( [matcher, subtype])])
        writer.writerow( ['Total number of Users', totalUsers])
        writer.writerow( ['Total number of Fingers', totalFingers])
        writer.writerow( ['Total number of Enroll images', totalEnrollImages])
        writer.writerow( ['Total number of images for FA', totalFAimages])
        writer.writerow( ['Total number of images for FR', totalFRimages])
        writer.writerow( ['Finger FTER', fingerFTER])
        writer.writerow( ['User FTER', userFTER])
        writer.writerow( ['TemplateCreator FTAR', templateFTAR])
        writer.writerow( ['CrossMatcher FTAR', crossmatcherFTAR])
        writer.writerow( ['TotalFTAR', totalFTAR])
        if previous :
            writer.writerow( labels)
            rowKeys = ['1K', '5K', '10K', '25K', '30K', '40K', '50K', '75K', '100K', '150K', '200K', '250K', '300K', '400K', '500K', '750K', '1M', '2M', '3M', '5M']
            for rowKey in rowKeys :
                writer.writerow( outRows[rowKey])
        elif inResultTable and line1k :
            writer.writerow( labels)
            writer.writerow( row1k)
            writer.writerow( row5k)
            writer.writerow( row10k)
            writer.writerow( row25k)
            writer.writerow( row30k)
            writer.writerow( row40k)
            writer.writerow( row50k)
            writer.writerow( row75k)
            writer.writerow( row100k)
            writer.writerow( row150k)
            writer.writerow( row200k)
            writer.writerow( row250k)
            writer.writerow( row300k)
            writer.writerow( row400k)
            writer.writerow( row500k)
            writer.writerow( row750k)
            writer.writerow( row1m)
            writer.writerow( row2m)
            writer.writerow( row3m)
            writer.writerow( row5m)
        writer.writerow( ['FAR/FRR at Executed Threshold', threshold, fas, far+'%' + farnnkStr, frs, frr+'%'])
        writer.writerow( ['Transactional FRR at Executed Threshold', threshold, '', '', tfrs, tfrr+'%'])
        writer.writerow( ['Elapsed Time', elapsedTime])
        writer.writerow( ['Average Enroll Time', avgEnroll])
        writer.writerow( ['Max Enroll Time', maxEnroll])
        writer.writerow( ['Min Enroll Time', minEnroll])
        if avgFAMatch == '-1' :
            writer.writerow( ['Average Match Time', avgMatch])
            writer.writerow( ['Max Match Time', maxMatch])
            writer.writerow( ['Min Match Time', minMatch])
        else :
            writer.writerow( ['Average FA Match Time', avgFAMatch])
            writer.writerow( ['Max FA Match Time', maxFAMatch])
            writer.writerow( ['Min FA Match Time', minFAMatch])
            writer.writerow( ['Average FR Match Time', avgFRMatch])
            writer.writerow( ['Max FR Match Time', maxFRMatch])
            writer.writerow( ['Min FR Match Time', minFRMatch])
        if enrollImageMax > 0 :
            writer.writerow( ['Max Enrollment Images', enrollImageMax])
        writer.writerow( ['Histogram of Finger Sample Matched Index'])
        writer.writerow( ['Total Found', histTotal])
        writer.writerow( ['Min Found', histMin])
        writer.writerow( ['Max Found', histMax])
        writer.writerow( ['Max Possible', histPossible])
        writer.writerow( ['Average', histAvg])
        writer.writerow( ['Std Dev', histStdDev])
        writer.writerow( histTitles)
        for line in histTable :
            writer.writerow( line)
        writer.writerow( ['Update Templates', updTempl])
        writer.writerow( ['Max Template Updates', maxUpdates])
        if opticalFound :
            writer.writerow( ['Max Primary Images', maxPrimaryImages])
            writer.writerow( ['Start Secondary At Next Image', startSecondary])
    FFAT_watermark_report.watermark_report( csvName)
    inFile = open( os.path.splitext( csvName)[0] + '_watermarked' + '.csv', newline='')
    reader = csv.reader( inFile)
    for i in range(9):
        inLine = next(reader)
        if i == 4 : continue
        if xlsx:
            if i == 0 or i == 5 :
                for col,cell in zip( range(len(inLine)), inLine):
                    allSheet.write( rowNum, col, cell)
            else :
                allSheet.write( rowNum, 0, int(inLine[0]))
                allSheet.write( rowNum, 1, int(inLine[1]))
                allSheet.write( rowNum, 2, float(inLine[2]))
                allSheet.write( rowNum, 3, inLine[3])
            rowNum += 1
        else:
            if i == 4 : continue
            writer.writerow( inLine)
    if xlsx:
        workbook.close()
    return 0

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Read an FFAT output CSV file from a run of FFAT and generate a summary report in CSV or Excel format.',
        epilog='Excel output requires XlsxWriter installed with Python.  Install from Confluence "Software / Development Environment / Python" into \\Python27\\Lib\\site-packages.')
    parser.add_argument("csvName", help='Name of the FFAT CSV file')
    parser.add_argument( "logName", help="Name of the FFAT log file.")
    parser.add_argument( '--add1ToTableVals', action='store_true', help='Compensate for FFAT bug by adding 1 to scores in 1:nnnK table')
    parser.add_argument( '--addFTA2FRR', action='store_true', help='Add Crossmatcher FTAR to the FRR values in the 1:nnnK table.')
    outGrp = parser.add_mutually_exclusive_group()
    outGrp.add_argument( '--xlsx', '-x', action='store_true', default=False, help='Output the report into an Excel file.')
    outGrp.add_argument( '--csv', '-c', action='store_true', default=True, help='Output the report into a CSV file.')
    args = parser.parse_args( )

    exit( summary_report( args.csvName, args.logName, args.add1ToTableVals, args.addFTA2FRR, args.xlsx ))
