# analysisEvtx
 
分析windows日志文件（.evtx），通过写定的xml tag值，将对应的日志内容转换为excell，方便进行数据筛选.   
开发语言：python   
主要使用python库：python-evtx    
写定的xml tag值："Provider.Name", "Provider.Guid", "EventID", "Level", "TimeCreated.SystemTime",    
                 "EventRecordID", "Execution.ProcessID", "Execution.ThreadID", "Channel",    
                 "Computer", "ProcessID", "Application", "Direction", "SourceAddress", "SourcePort",    
                 "DestAddress", "DestPort", "Protocol", "RemoteUserID", "RemoteMachineID",    
                 "Security.UserID", "QueryName", "EventSourceName", "Data"    
                 如需修改则修改main.py中__init__函数的requireTagList数组    
