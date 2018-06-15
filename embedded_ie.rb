

  $VERBOSE=nil
  require 'watir/win32ole/win32ole.so' # This might throw some constant initialization warnings 
  $VERBOSE=false
  require 'windows/com'
  require 'windows/unicode'
  require 'windows/error'
  require 'windows/national'
  require 'windows/window/message'
  require 'windows/msvcrt/buffer'

  include Windows::COM
  include Windows::Unicode
  include Windows::National
  include Windows::Error
  include Windows::Window::Message
  include Windows::MSVCRT::Buffer
  
##################################################################
# Global Variables  
##################################################################

	# None

##################################################################
# Constants 
##################################################################

  # None

##################################################################
# Modules/Classes
##################################################################

  module Watir
        
    class IE
      
      # Method to attach to an embedded browser Ex: Watir::IE.attach_embedded("Main Window", :hnd, 0x005F0658) or  Watir::IE.attach_embedded("Main Window", :classnameNN, "Internet Explorer_Server1")  
      def self.attach_embedded(parent, how, what) 
        ie = new true
        ie.attach_eie(parent, how, what)
        ie  
      end
      
      # Method to attach to an embedded browser based on handle/class name of the control
      def attach_eie(parent, how, what)           
        case how.to_s        
            when "hnd"
                handle = what
            when "classnameNN"
                handle = Watir.autoit.ControlGetHandle(parent, "", what)              
        end            
        iwebbrowser2_com_object = get_webbrowser2(handle)     # get web browser control from ie handle          
        if iwebbrowser2_com_object == nil
          raise NoMatchingWindowFoundException,
                 "Unable to locate a window with #{how} of #{what}"
        end         
        @ie =  iwebbrowser2_com_object
        initialize_options
        wait
      end
      
      # Method to get the embedded browser COM onject (IWebBrowser2-Automation Interface)
      def IWebBrowser2_object()
        @ie
      end
      
    end

  end

##################################################################
# Functions
##################################################################

  SMTO_ABORTIFHUNG = 0x0002
  ObjectFromLresult = Win32::API.new('ObjectFromLresult', 'LPIP', 'L', 'oleacc')
  IID_NULL = [0x00000000,0x0000,0x0000,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00].pack('LSSC8')

  #***********************************************************************************************************
  # DESCRIPTION: Function to to retrieve IHTMLDocument2 COM object from the 'Internet Explorer_Server'
  #                         control/embedded ie window handle
  # INPUT: Embedded ie handle
  # OUTPUT: IHTMLDocument2 COM object
  # EXAMPLE OF USE: get_htmldocument2(0x005F0658)
  #***********************************************************************************************************   
  def get_htmldocument2(hnd)

    CoInitialize(0)
    reg_msg = RegisterWindowMessage("WM_HTML_GETOBJECT")
    iid2 =[0x332C4425,0x26CB,0x11D0,0xB4,0x83,0x00,0xC0,0x4F,0xD9,0x01,0x19].pack('LSSC8')
    result = 0.chr*4
    SendMessageTimeout(hnd.hex, reg_msg, 0, 0, SMTO_ABORTIFHUNG,1000, result)
    result = result.unpack('L')[0]
    pDisp = 0.chr * 4
    r = ObjectFromLresult.call(result, iid2, 0, pDisp)
    if r == 0
      pDisp = pDisp.unpack('L').first
      ucPtr = multi_to_wide("parentWindow")
      lpVtbl = 0.chr * 4
      table = 0.chr * 28
      memcpy(lpVtbl,pDisp,4)
      memcpy(table,lpVtbl.unpack('L').first,28)
      table = table.unpack('L*')
      getIDsOfNames = Win32::API::Function.new(table[5],'PPPLLP','L')
      dispID = 0.chr * 4
      getIDsOfNames.call(pDisp,IID_NULL,[ucPtr].pack('P'),1,LOCALE_USER_DEFAULT,dispID)
      dispID = dispID.unpack('L').first
      dispParams = [0,0,0,0].pack('LLLL')
      res = 0.chr * 16
      invoke = Win32::API::Function.new(table[6],'PLPLLPPPP','L')
      hr = invoke.call(pDisp,dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT,DISPATCH_PROPERTYGET | DISPATCH_METHOD,dispParams, res, nil, nil)
      if hr != S_OK
        raise StandardError, "IDispatch::Invoke() failed with %08x" % hr
      end
      res = res.unpack('SSSSLL')
      if res[0] == VT_DISPATCH
        return res[4]
      end
    else
      nil
    end
  end
  
  #***********************************************************************************************************
  # DESCRIPTION: Function to retrieve IWebBrowser2 COM object from the 'Internet Explorer_Server' control/embedded ie window handle
  # INPUT:  
  # OUTPUT: IWebBrowser2 COM object
  # EXAMPLE OF USE: get_webbrowser2(0x005F0658)
  #***********************************************************************************************************	 
  def get_webbrowser2(hnd)

    iid_IServiceProvider = [0x6d5140c1,0x7436,0x11ce,0x80,0x34,0x00,0xaa,0x00,0x60,0x09,0xfa].pack('LSSC8')
    iid_IWebBrowserApp = [0x0002DF05,0x0000,0x0000,0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46].pack('LSSC8')
    iid_IWebBrowser2 = [0xD30C1661,0xCDAF,0x11D0,0x8A,0x3E,0x00,0xC0,0x4F,0xC9,0xE2,0x6E].pack('LSSC8')

    pDisp = get_htmldocument2(hnd)
    return nil if pDisp.nil?

    lpVtbl = 0.chr * 4
    table = 0.chr * 28
    memcpy(lpVtbl,pDisp,4)
    memcpy(table,lpVtbl.unpack('L').first,12)
    table = table.unpack('L*')
    queryInterface = Win32::API::Function.new(table[0],'PPP','L')
    dispID = 0.chr * 4
    queryInterface.call(pDisp,iid_IServiceProvider,dispID)
    pDisp = dispID.unpack('L').first

    lpVtbl = 0.chr * 4
    table = 0.chr * 28
    memcpy(lpVtbl,pDisp,4)
    memcpy(table,lpVtbl.unpack('L').first,16)
    table = table.unpack('L*')
    queryService = Win32::API::Function.new(table[3],'PPPP','L')
    dispID = 0.chr * 4
    queryService.call(pDisp,iid_IWebBrowserApp,iid_IWebBrowser2,dispID)
    dispID = dispID.unpack('L').first
    if dispID>0
      WIN32OLE.connect_unknown(dispID)
    else
      nil
  end
  end
  
##################################################################

# Usage:

# require 'watir'
# require 'embedded_ie'

# embedded_browser = Watir::IE.attach_embedded("Parent class", :classnameNN, "Internet Explorer_Server1")
# puts embedded_browser.url
 
you can use watir methods ...



