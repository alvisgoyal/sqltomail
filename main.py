if __name__ =="__main__":
    import Init
    import os
    import threading
    
    list_thread =[]
    
    Init.InitialMailData()
    Init.InitialSQL()
    
    curr_thread=0
    while curr_thread < no_sqlQuery :
        
        list_thread[curr_thread].start()
        curr_thread+=1
    
    
    curr_thread=0
    while curr_thread < no_sqlQuery :
        
        list_thread[curr_thread].join()
        curr_thread+=1