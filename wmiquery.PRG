**************************************************************
* Marco Plaza, 2018
* @nfoxProject
* https://github.com/nftools/wmiquery
**************************************************************
*
* WMI Query Tool : returns an object with item count & items array from wmiquery
*
* simple usage: wmiQuery( wmiQuery [, wmiClass] ) 
* ( wmiclass defaults to "CIMV2" )
*
* Returns: Object 
*
* Object schema ( as nxs https://github.com/nftools/nxs ):
*
*.oWmiResult:
*      .count = i
*      .items[]
*         -item = v
*
* sample procedure "testme" included below
*
*****************************************************************
Parameters wmiquery,wmiclass

Private All

wmiclass = Evl(m.wmiclass,'CIMV2')
wmiquery = Evl(m.wmiquery,'')

emessage = ''

Try
   objwmiservice = Getobject("winmgmts:\\.\root\"+m.wmiclass)
   oquery = objwmiservice.execquery( 'SELECT * FROM '+m.wmiquery,,48)
   owmi = processobject( oquery )
Catch To oerr
   emessage = m.oerr.Message
Endtry

If !Empty(m.emessage)
   Error ' Invalid Query/WmiClass '
   Return .Null.
Else
   Return m.owmi
Endif

*-------------------------------------------------
Procedure processobject( oquery )
*-------------------------------------------------
Private All

owmi = Createobject('empty')
AddProperty(owmi,'items(1)',.Null.)
nitem = 0

Try

   For Each oitem In m.oquery

      nitem = m.nitem + 1
      Dimension owmi.items(m.nitem)
      owmi.items(m.nitem) = Createobject('empty')
      setproperties( m.oitem, owmi.items(m.nitem) )

   Endfor

Catch

Endtry

AddProperty(owmi,'count',m.nitem)

Return m.owmi

*--------------------------------------------------------
Procedure setproperties( oitem , otarget  )
*--------------------------------------------------------
Private All

For Each property In m.oitem.properties_
   Try
      Do Case
      Case Vartype( m.property.Value ) = 'O'
         thisproperty = Createobject('empty')
         setproperties(m.property.Value, m.thisproperty )
         AddProperty( otarget ,m.property.Name,m.thisproperty)

      Case m.property.isarray

         AddProperty( otarget ,property.Name+'(1)',.Null.)
         thisarray = 'otarget.'+m.property.Name

         nitem = 0

         If !Isnull(m.property.Value)

            For Each Item In m.property.Value

               nitem = m.nitem+1
               Dimension &thisarray(m.nitem)

               If Vartype( m.item) = 'O'
                  thisitem = Createobject('empty')
                  setproperties( m.item, m.thisitem )
                  &thisarray(m.nitem) = m.thisitem
               Else
                  &thisarray(m.nitem) = m.item
               Endif

            Endfor

         Endif

      Otherwise
         AddProperty( otarget ,m.property.Name,m.property.Value)
      Endcase

   Catch To oerr
      Messagebox(Message(),0)
   Endtry
Endfor

*---------------------------------------------------------------------------------
Procedure testme 
* note:
* this code uses underscore ( _.prg ) as modern replacement for addproperty()
* available at https://raw.githubusercontent.com/nftools/underscore/master/_.prg
* the "addobject" version is below ( function testme_no_ )
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

*----------------------------------
procedure testme_no_
*----------------------------------
Public oinfo

oinfo = Create('empty')

Wait 'Running WMI Query....please wait.. ' Window Nowait At Wrows()/2,Wcols()/2


addproperty( oinfo, "monitors"  , wmiquery('Win32_PNPEntity where service = "monitor"') )
addproperty( oInfo, "diskdrive" , wmiquery('Win32_diskDrive') )
addproperty( oInfo, "startup" ,   wmiquery('Win32_startupCommand'))
addproperty( oInfo, "BaseBoard" , wmiquery('Win32_baseBoard') )
addproperty( oInfo, "netAdaptersConfig",  wmiquery('Win32_NetworkAdapterConfiguration') )


Messagebox( 'Please explore "oInfo" in debugger watch window or command line ',0)


