# -*- coding:gbk -*-


def printStruct(jceStruct, logger, level = 0):
    result = printStructToString(jceStruct, level + 1)
    logger.debug("\n%s\n%s", jceStruct.__class__, result)

def printItemToString(item, field, level = 0):
    logresult = ""
    if isinstance(field, int):
        logresult += "%s%s : %d\n" % (level * '\t', item, field)
    elif isinstance(field, str):
        logresult += "%s%s : %s\n" % (level * '\t', item, field)
    elif isinstance(field, list):
        logresult += "%s%s : size:%d\n" % (level * '\t', item, len(field))
        listcount = 0
        for listitem in field:
            listcount += 1
            logresult += "%s%s : %d\n" % ((level + 1) * '\t', 'item', listcount)
            try:
                logresult += printItemToString("", listitem, level + 2)
            except:
                logresult += 'unsupport type: %s\n' % type(field)
    elif isinstance(field, dict):
        logresult += "%s%s : size:%d\n" % (level * '\t', item, len(field))
        dictcount = 0
        keys = sorted(field.keys())
        for dictkey in keys:
            dictcount += 1
            logresult += "%s key : %s\n" % ((level + 1) * '\t', str(dictkey))
            try:
                logresult += printItemToString("", field[dictkey], level + 2)
            except:
                logresult += 'unsupport type: %s\n' % type(field)
    else:
        try:
            logresult += "%s%s : %s\n" % (level * '\t', item, field.__class__)
            logresult += printStructToString(field, level + 1)
        except:
            logresult += 'unsupport type: %s\n' % type(field)

    return logresult

def printStructToString(jceStruct, level = 0):

    logresult = ""
    fields = vars(jceStruct)
    keys = sorted(fields.keys())
    for item in keys:
        logresult += printItemToString(item, fields[item], level)

    return logresult
