# WMIQUERY Function

Allows you to make synchronous Wmi Querys over local PC and get result as a VFP object 
using a simple function call.

     wmiQuery( wmiQuery [, wmiClass] )
 
### Parameters

**wmiQuery**

Any valid WMI Query 

**wmiClass**

Specify the wmiClass name; defaults to "CIMV2"

### Return Value: 

Object. 


	Result object schema ( nxs format https://github.com/nftools/nxs ): 

    .oWmiResult:
      .count = i
      .items[]
        -item = v
    
### Sample procedure ( included in wmiquery.prg ): 

    *---------------------------------------------------------------------------------
    Procedure testme 
    * note:
    * this code uses underscore ( _.prg ) as modern replacement for addproperty()
    * available at https://raw.githubusercontent.com/nftools/underscore/master/_.prg
    * a regular version ( "testme_no_" ) using "addobject" is also included 
    *---------------------------------------------------------------------------------
    Public oinfo
    
    oinfo = Create('empty')
    
    Wait 'Running WMI Query....please wait.. ' Window Nowait At Wrows()/2,Wcols()/2
    
    
    With _( m.oinfo )
       .monitors  =  wmiquery('Win32_PNPEntity where service = "monitor"')
       .diskdrive =  wmiquery('Win32_diskDrive')
       .startup   =  wmiquery('Win32_startupCommand')
       .BaseBoard =  wmiquery('Win32_baseBoard') 
       .netAdaptersConfig = wmiquery('Win32_NetworkAdapterConfiguration')
    Endwith
    
    
    Messagebox( 'Please explore "oInfo" in debugger watch window or command line ',0)
    


![](https://github.com/nftools/wmiQuery/blob/master/wmiquery.jpg)

