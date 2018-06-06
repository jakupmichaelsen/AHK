#SingleInstance force
#NoTrayIcon
#Include, D:\Dropbox\Code\AHK\#Include\HotstringHelper.ahk
; Menu, Tray, Icon, Letter-T.ico

; Folders
:*::db::D:\Dropbox{Enter}
:*::fc::C:\Users\Vostro3360\AppData\Local\Temp\fcctemp{Enter}
:*::fevernote::C:\Users\Vostro3360\AppData\Local\Evernote\Evernote\Databases\Attachments{Enter}
:*::fadm::C:\Users\Vostro3360\Google Drev\Adjunkt\Administration{Enter}
:*::grammatik::C:\Users\Vostro3360\Google Drev\Adjunkt\Grammatik{Enter}
:*::units::C:\Users\Vostro3360\Google Drev\Adjunkt\Teaching Units{Enter}

; MAIL TO WEB APPS
:*:2.0class@trello::jakupmichaelsen+uyfhz59vj5obdxenyout@boards.trello.com
:*:todo@trello::jakupmichaelsen+rccx0d4km3ltdxcjsigu@boards.trello.com
:*:class@trello::jakupmichaelsen+68lya8uxhdcome77lgxe@boards.trello.com
:*:mailtoevernote::jakupmichaelsen.be37f@m.evernote.com

; Mobile network mode
::pwcdma::networkType:0 ; WCDMA Preferred
::ogsm::networkType:1 ; GSM only <-- This would be "2G" on GSM networks
::owcdm::networkType:2 ; WCDMA only <--WCDMA is "3G" on GSM networks. You may know it as HSPA
::agsm.::networkType:3 ; GSM auto (PRL)
::acdma::networkType:4 ; CDMA auto (PRL)
::ocdma::networkType:5 ; CDMA only <-- This would be "2G" on CDMA networks
::oevdo::networkType:6 ; EvDo only <-- EvDo is "3G" on CDMA networks
::agsmcdma::networkType:7 ; GSM/CDMA auto (PRL)
::altecdma::networkType:8 ; LTE/CDMA auto (PRL)
::altegsm.::networkType:9 ; LTE/GSM auto (PRL)
::altegsmcdma::networkType:10 ; LTE/GSM/CDMA auto (PRL)
::olte::networkType:11 ; LTE only
 
; #IfWinActive, D:\Dropbox\Code\AHK\TextExpansion.ahk (Code) - Sublime Text (UNREGISTERED)
+F12:: ; Format for Xposed Macro Expand, save & export, and reset formatting 
	Sleep 50
	Send {Esc}
	Sleep 50
	Send ^h
	Sleep 50
	Send :*:
	Send {Tab}
	Send ^a{BackSpace}
	Send ::
	Sleep 50
	Send ^!{Enter}
	Send {Esc}
	Send ^s
	Sleep 100
	FileCopy, D:\Dropbox\Code\AHK\TextExpansion.ahk, D:\Dropbox\Code\AHK\AHKimport.ahk, 1
	Sleep 100
	Send ^z 
	Sleep 50
	Send ^s
Return

; Text 
:*:ggss::Google Sheets
:*::elevmails::15h_Bg6sj8aLkqS2j_DTMibhdCsYcsyhJm3j_hjgv5Kc
:*:kbsc::keyboard shortcut
:*:jmg::jakupmichaelsen@gmail.com
:*c:jjmm::jakupmichaelsen
:*c:JJMM::Jákup Michaelsen
:*:kktt::kristinathastrup@hotmail.com
:*:unilogin::jaku0288
:*:engadm::143957-Herning HF & VUC
:*:sidk::site:dk 
:*:siuk::site:uk 
:*:fpdf::filetype:pdf
:*:fppt::filetype:ppt
:*:fdoc::filetype:doc
:*:gtda::=GOOGLETRANSLATE(a1,"en","da")
:*:gten::=GOOGLETRANSLATE(a1,"da","en")
::gradelist::https://docs.google.com/spreadsheets/d/1T2k2JE_cHT2crm-Vu3v6_HA964g-_v-_rYN4M0qNyk4
:*:drivedl::https://drive.google.com/uc?export=download&id= ; <- + file ID
:*:moodlepwd::jdm060814
:*:engrampwd::38920140812105808
:*:tv2playpwd::2F75JrCS
:*:uploadtodrive::kortlink.dk/eqyq
:*:ssgg::Silkeborg Gymnasium
:*:trs::08:05
:*:trf::15:25
:*:tr2f::14:05
:*:regexAB::[\S\s]*?
:*:aahk::autohotkey
:*:rgi::rigtig god idé 
:*:lmpe::læs mere på ENGRAM
:*:mvh::Med venlig hilsen
:*:ofl::Overflødigt{Esc}
:*:loremipsum::Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.
:*:sstt::Sublime Text
:*:\1pr�::!nc{Enter}Limit each paragraph to one main idea. You have too many threads of thought in this paragraph.{Esc}