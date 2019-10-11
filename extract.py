# Use orderedDict to keep the semantic order of the substring

import olefile,string
from collections import OrderedDict,Counter
import sys, unicodedata, re
import base64

list_objects = ["Macros","WordDocument","Data","1Table"]
susp_obj = ''
max_size = 0
ocurrences = 0
dif = 0
substr = ''
substr2 = ''
subj = 1
total = 0
result = OrderedDict()
ss = ''
h=0
last=''

def dict_an(s):
        res = OrderedDict()
        #result = {'A':0}
        for letter in s:
                if letter in res:
                        res[letter] += 1
                else:
                        res[letter] = 1
        return res


with olefile.OleFileIO('attach1.doc') as ole:
        if ole.exists('worddocument'):
                print("This is a Word document.")
                if ole.exists('macros/vba'):
                        print("This document seems to contain VBA macros.")
                        for i in ole.listdir():
                                if i[0] not in list_objects:
                                        print ("Obj: %s \t Size: %5d" % ('/'.join(i),ole.get_size(i)))
                                        if max_size < ole.get_size(i):
                                                susp_obj = i
                                                max_size = ole.get_size(i)
                                        else:
                                                continue
                        print " ========================================================================================"
                        print " [DOC] - Suspicious object:"
                        print " ========================================================================================"
                        print (" OBJECT: \t %s" % susp_obj)
                        print (" SIZE: \t %5d" % max_size)
                        data = ole.openstream(susp_obj).read()

                        data = data.replace('\$','')
                        data = ''.join(data.split())
                        data = data[25:-28]
                        #data = data[]
                        dd = dict_an(data)
                        print ("Datos: %s" % (data))

                        # # -----------------------------------------------------------------------------------------------
                        for (i) in dd.keys():
                               print "key: %s\t value: %5d" % (i,dd[i])
                        d = Counter(dd)
                        top10 = d.most_common(10)
                        for (j,k) in top10:
                               print "Key: %s Value: %5d" % (j,k)
                        subj = int(top10[0][1])
                        for (i,j) in top10:

                                if (subj-j) < 300:
                                    substr += i
                                else:
                                    break
                                subj = j
                        print "Substr: %s" % substr
                        for k in dd.keys():
                                if k in substr:
                                    result[k] = dd.keys().index(k)

                        for i in range(1,(len(result.values()))):
                            if (result.values()[i] - result.values()[i-1]) > 1:
                                last = 'n'
                            else:
                                ss += result.keys()[i-1]
                                last = 'y'

                        if last == 'y':
                                ss += result.keys()[i]
                        else:
                                pass


                        data = ''.join([str(char) for char in data if char in string.printable])

                        higherValue = 0 
                        max = 0
                        
                        print "SS: %s" % ss
                        for key in range(len(ss)-1):
                            #if key == len(ss): break
                            current = len(data.split(ss[key:]))
                            if current > higherValue and higherValue == 0: higherValue = current
                            elif current > higherValue:
                                ss = ss[key:]
                                break
                        
                        print "SUBSTRING: %s" % ss
                        resultado = data.replace(ss,'')

        ole.close()

        script = base64.b64decode(resultado)

        print " ========================================================================================"
        print " [SCRIPT] Script en B64:"
        print " ========================================================================================"
        print  resultado
        print " ========================================================================================"
        print " [OBFUSCATION SUBSTRING DETECTED]: \t <<  %s  >> " % ss
        print " ========================================================================================"
        print " ========================================================================================"
        print " [DECODED SCRIPT] "
        print " ========================================================================================"
        print ''.join([str(char) for char in script if char in string.printable]) 
        print #script.replace(';',';\n')
        print " =======================================================================================\n"

