contrList = '@Contracts2015.dif'
statFile = 'stat_1.dif'
month = 'January'

def readcontr(filename):
    f=open(filename, 'r')
    contracts = {}
    for line in f:
        contracts[line.split(';')[0]] = [line.split(';')[1],line.split(';')[2],line.split(';')[3]]
    f.close()
    print '---> contract names have been loaded'
    return contracts

def readamount(filename, contracts):
    f=open(filename, 'r')
    for line in f:
        if line.split(';')[0] in contracts:
            contracts[line.split(';')[0]].append(line.split(';')[1][0:-1])
    f.close()
    print '---> contract amounts have been loaded:'
    print contracts
    return contracts

def createmsg(a,b,month):
    cid = a
    if len(b)<>4:
        print '(!) Error: There is no data in the bill for contract', b
        return
    name,surname,typ,amount = b
    if 'N' in typ: return
    import win32com.client
    from win32com.client import constants as const
    o = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    NS = o.GetNamespace('MAPI')

    TmpAdmin = """<HTML><BODY><P>Dear %s,<BR><BR>
    This email contains detailed monthly bill for MTS Ukraine mobile service usage (attached). <BR>
    This is for your information.<BR><BR> Amount: %s UAH</P><BR>
    """

    Msg = o.CreateItem(0)
    Msg.BodyFormat = 1
    if 'A' in typ:
        Msg.HTMLBody = TmpAdmin % (name,amount)
        Msg.Subject = str("FYI: Mobile Phone Bill for "+month)
    Msg.To = str(name+'.'+surname+'@domain.com')
    print str(cid+'.xls')
    attachment1 = str('C:\\mtsscript\\'+cid+'.xls')
    Msg.Attachments.Add(attachment1)
    Msg.Save()

#init the contracts list:
contracts = readcontr(contrList)
readamount(statFile, contracts)
for a,b in contracts.iteritems():
    createmsg(a,b,month)
