import ConfigParser
def getProperty(sectionName,key):
    # Return a property for the given key
    config = ConfigParser.ConfigParser()
    config.readfp(open('C:\MailboxMonitoring\Configurations\ConfigFile.properties'))
    return config.get(sectionName, key)
