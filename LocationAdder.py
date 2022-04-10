import sys



def NextLocation(LOCATION):
        ''' returns string LOCATION after adding 1 to the calling parameter'''
        itemchar = LOCATION[4:]
        binchar = LOCATION[3:4]
        traychar = LOCATION[2:3]
        shelfchar = LOCATION[1:2]
        rackchar = LOCATION[0:1]


        carrytobins = False
        carrytotrays = False
        carrytoshelves = False
        carrytorack = False
        interimlocation = rackchar + shelfchar + traychar + binchar  + itemchar

        if itemchar == '05':
            itemchar ='01'
            carrytobins = True
        else:
            i = ord(itemchar[1])
            i = i + 1
            itemchar = '0' + chr(i)

        if carrytobins and binchar == 'o':
            binchar = 'a'
            carrytotrays = True
        elif carrytobins:
            j = ord(binchar[0])
            j = j + 1
            binchar = chr(j)

        if carrytotrays and traychar  == '4':
            carrytoshelves = True
            traychar = '1'
        elif carrytotrays:
            k = ord(traychar[0])
            k = k + 1
            traychar = chr(k)

        if carrytoshelves and shelfchar == 't':
            carrytorack = True
            shelfchar = 'a'
        elif carrytoshelves:
            l = ord(shelfchar[0])
            l = l + 1
            shelfchar = chr(l)
            print('Shelf is: ', shelfchar)
        
        if rackchar == '9' and carrytorack:
            print('Storage is full - No location ID available')
        elif carrytorack:
            m = ord(rackchar[0])
            m = m  + 1
            rackchar = chr(m)

        LOCATION = rackchar + shelfchar + traychar + binchar + itemchar
        return(LOCATION)

if __name__ == '__main__':
        CURRENT_VERSION = argv[0][:-3]
        COPYRIGHT_VERSION_NOTICE = '<!-- ' + CURRENT_VERSION +  'Copyright 2022 myearringdepot.com All Rights Reserved  -->'    
        LOCATION = 'Aa1C05'
        print(COPYRIGHT_VERSION_NOTICE)
        print(f'Last LOCATION =: {LOCATION}, first new LOCATION will be:  {NextLocation(LOCATION)}')