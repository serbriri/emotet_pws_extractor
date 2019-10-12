from collections import OrderedDict, Counter
import olefile, string, base64
import sys, unicodedata, re
import argparse

def buildStats(s):
    res = OrderedDict()
    for letter in s:
        if letter in res:
            res[letter] += 1
        else:
            res[letter] = 1
    return res

def extractSubstring(data):
    
    sub1 = ''
    subj = 0
    last = ''
    ss = ''
    charStatsDict = OrderedDict()
    dd = OrderedDict()

    dd = buildStats(data)
    d = Counter(dd)
    top10 = d.most_common(10)
    subj = int(top10[0][1])
    for (i, j) in top10:
        if (subj - j) < 300:
            sub1 += i
        else:
            break
        subj = j
    for k in dd.keys():
        if k in sub1:
            charStatsDict[k] = dd.keys().index(k)

    for i in range(1, (len(charStatsDict.values()))):
        if (charStatsDict.values()[i] - charStatsDict.values()[i - 1]) > 1:
            last = 'n'
        else:
            ss += charStatsDict.keys()[i - 1]
            last = 'y'

    if last == 'y':
        ss += charStatsDict.keys()[i]
    else:
        pass     
    data = ''.join([str(char) for char in data if char in string.printable])
    higherValue = 0
    for key in range(len(ss) - 1):
        current = len(data.split(ss[key:]))
        if current > higherValue and higherValue == 0:
            higherValue = current
        elif current > higherValue:
            ss = ss[key:]
            break
    return ss

def main():
    list_objects = ["Macros", "WordDocument", "Data", "1Table"]
    finalScript = ''
    susp_obj = ''
    subStr = ''
    max_size = 0
    ocurrences = 0
    dif = 0
    subj = 1
    total = 0
    h = 0
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help=".doc file")
    parser.add_argument("-s", "--script", help="Output only decoded script", action="store_true")
    parser.add_argument("-a", "--all", help="Get all details", action="store_true")
    parser.add_argument("-sub", "--substring", help="print only obfuscation string", action="store_true")
    args = parser.parse_args()
    try:
        with olefile.OleFileIO(args.file) as ole:
            if ole.exists('worddocument'):
                if ole.exists('macros/vba'):
                    for i in ole.listdir():
                        if i[0] not in list_objects:
                            if max_size < ole.get_size(i):
                                susp_obj = i
                                max_size = ole.get_size(i)
                            else:
                                continue
                    data = ole.openstream(susp_obj).read()
                    data = data.replace('\$', '')
                    data = ''.join(data.split())
                    data = data[25:-28]
                    subStr = extractSubstring(data)
                    finalScript = data.replace(subStr, '')
                    script = base64.b64decode(finalScript)
    except:
        print("[ ERROR ]: File does not exists or is not an OLE file.")
        exit(0)

    if args.all:
        print " [Obfuscation string]: %s " % subStr
        print " [Decoded script]:\n %s " % script.replace(';',';\n')
        
    if args.script:
        print script.replace(';',';\n')
    
    if args.substring:
        print subStr  
    
if __name__ == "__main__": 
    main()
