from collections import OrderedDict, Counter
import olefile, string, base64
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
    subj = 0
    last = ''
    rep_ind = [0] * 10
    resss = [0] * 5
    count = 0
    def_count = 0
    cont = True
    charStatsDict = OrderedDict()

    d = Counter(buildStats(data))
    top10 = d.most_common(10)  # Extraer los 10 caracteres mas repetidos
    last = 'y'
    if (top10[0][1] - top10[((len(top10)/2) - 1)][1]) > 1500:
        for i in range(len(top10)):
            subj = top10[i + 1][1]
            # Controlar si se repite el caractaer o no basandose en su frecuencia
            dif_s = top10[i][1] - subj

            if dif_s < 50 and cont:  # Se repite y forma parte de la substring
                rep_ind[i] = 1
                resss[i] = top10[i][0]
                count += 1
                last = 'y'
            else:
                count += 1
                cont = False
                resss[i] = top10[i][0]
                if (top10[0][1] / 2) - subj > 300:  # Si la distancia con el caracter de substrings que no se repite es mayor de 300 consideramos
                    # que ya no pertenece al substring
                    break
                else:  # rep_ind[i] = 1
                    if last == 'y':
                        rep_ind[i - 1] = 1
                        rep_ind[i] = 0

        rep_char = [''] * count
        def_count = count  # contador con duplicados

        for i in range(count):
            if rep_ind[i] == 1:
                rep_char[i] = resss[i]
                def_count += 1

    subss = ''
    data = ''.join([str(char) for char in data if char in string.printable])

    for char in data:
        if (char in resss) and (def_count != 0):
            def_count -= 1
            subss += char

    return subss

def main():
    list_objects = ["Macros", "WordDocument", "Data", "1Table"]
    finalScript = ''
    script = ''
    susp_obj = ''
    subStr = ''
    max_size = 0

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
                    data = data[25:-28]
                    data = ''.join([str(char) for char in data if char in string.printable])
                    subStr = extractSubstring(data)
                    print(subStr)
                    finalScript = data.replace(subStr, '')
                    print(finalScript)
                    script = base64.b64decode(finalScript)
                    print(script)
    except:
        print("[ ERROR ]: File does not exists or is not an OLE file.")
        exit(0)

    if args.all:
        print(" [Obfuscation string]: %s " % subStr)
        print(" [Decoded script]:\n %s " % script.replace(';', ';\n'))

    if args.script:
        print(script.replace(';', ';\n'))

    if args.substring:
        print(" [Obfuscation string]: %s " % subStr)

if __name__ == "__main__":
    main()
    
