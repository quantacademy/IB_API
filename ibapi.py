"""
This script will contain functions for accessing the IB API
"""

#===============================================================================
# LIBRARIES
#===============================================================================
from ib.opt import ibConnection, message
from time import sleep

#===============================================================================
# Class IB_API
#===============================================================================
class IB_API:
    """
    This class will establish a connection to IB and group the different 
    operations
    """
    def __init__(self):
        """
        Connection to the IB API
        """
        print "Calling connection"
        # Creation of Connection class
        self.connection = ibConnection()
        # Register data handlers
        self.connection.registerAll(self.process_messages)
        # Connect
        self.connection.connect()
        
    def process_messages(self, msg):
        """
        Function that indicates how to process each different message
        """
        if msg.typeName == "updatePortfolio":
            print msg
            
    def get_account_updates(self):
        """
        Call for updated portfolio information
        """
        print "Calling Portfolio"
        self.connection.reqAccountUpdates(1, '')
        sleep(10)
    
 
if __name__ == '__main__':
    ib = IB_API()
    ib.get_account_updates()
    
 