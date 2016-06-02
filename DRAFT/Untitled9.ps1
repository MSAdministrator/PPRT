
#check and see if Database has previously been created
    #Check the registry for information about the database
        #If it has been created, then open connection
#if it has not been created
    #then create the database
        Create-PPRTDatabase -Server $Serverz
        #then create the Tables
            Create-PPRTDatabaseTables -Server $Server
            #add this information to the Registry


$cmd.commandtext = "INSERT INTO servers (servername,username,spversion,reason) VALUES('{0}','{1}','{2}','{3}')" -f $os.__SERVER,$env.username,$os.servicepackmajorversion,$reason
$cmd.executenonquery()
$conn.close()