#! python3
# Password manager using command line argument.

password={'****':'******','*****':'*******','******': '*****','*****':'*******'}

import pyperclip, sys


if len(sys.argv)<2:
    print('Usage: python pw.py [account] - copy account password')
    sys.exit()

account=sys.argv[1]

if account in password:
    pyperclip.copy(password[account])
    print('Password for '+account+' copied to clipboard.')
else:
    print('There is no account named '+account)
    print('Do you want to update password dictionary?(y/n): ')
    yesno=input()
    if yesno.lower()=='y':
        print('Please input the password: ')
        accountPassword=input()
        fileObj=open('C:\\Users\\rashi\\python\\passwordManager.py')
        x=fileObj.read()
        fileObj.close()
        c=0
        for i in x:
            if i!='}':
                c+=1
            else:
                break
        x=x[:c]+",'"+account+"':'"+accountPassword+"'"+x[c:]
        fileObj=open('C:\\Users\\rashi\\python\\passwordManager.py', 'w')
        fileObj.write(x)
        fileObj.close()
        print('Updated.')

    else:
        print('Thank you Rashed!')

