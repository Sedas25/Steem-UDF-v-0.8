#include-once
#include <String.au3>
Global $hSession, $hConnect
Global Const $_JSONNull = Default
Global Const $WINHTTP_NO_REFERER = ""
Global Const $WINHTTP_NO_REQUEST_DATA = ""
Global Const $WINHTTP_NO_ADDITIONAL_HEADERS = ""
Global Const $WINHTTP_ACCESS_TYPE_NO_PROXY = 1
Global Const $WINHTTP_FLAG_SECURE = 0x00800000
Global Const $WINHTTP_FLAG_ESCAPE_DISABLE = 0x00000040
Global Const $INTERNET_DEFAULT_PORT = 0
Global Const $WINHTTP_NO_PROXY_NAME = ""
Global Const $WINHTTP_NO_PROXY_BYPASS = ""
Global Const $WINHTTP_QUERY_RAW_HEADERS_CRLF = 22
Global Const $WINHTTP_HEADER_NAME_BY_INDEX = ""
Global Const $WINHTTP_NO_HEADER_INDEX = 0
Global Const $WINHTTP_DEFAULT_ACCEPT_TYPES = 0
Global Const $hWINHTTPDLL__WINHTTP = DllOpen("winhttp.dll")
DllOpen("winhttp.dll")

; #CURRENT# =================================================================================
;_Steem_Start_Node
;_Steem_Close_Node
;_Steem_Get_Account_Votes
;_Steem_Get_Reward_Fund_Global
;_Steem_Get_Config
;_Steem_Get_State
;_Steem_Get_Dynamic_Global_Properties
;_Steem_Get_Chain_Properties
;_Steem_Get_Current_Median_History_Price
;_Steem_Get_Feed_History
;_Steem_Get_Witness_Schedule
;_Steem_Get_Hardfork_Version
;_Steem_Get_Next_Scheduled_Hardfork
;_Steem_Get_Account
;_Steem_Get_Account_Count
;_Steem_Get_Owner_History
;_Steem_Get_Block_Header
;_Steem_Get_Block
;_Steem_Get_Vesting_Delegations
;_Steem_Get_Ops_In_Block
;_Steem_Get_Content_Replies
;_Steem_Get_Active_Votes
;_Steem_Get_Content
; ===========================================================================================

; ===========================================================================================
;	_Steem_Start_Node($node)
;	$node = The ws / wss node too connect without the ws:// or wss://
; ===========================================================================================

Func _Steem_Start_Node($node)
	Global $hSession = _WinHttpOpen()
	If StringInStr($node, ":") Then
		$nodeSplit = StringSplit($node, ":")
		Global $hConnect = _WinHttpConnect($hSession, $nodeSplit[1], $nodeSplit[2])
	Else
		Global $hConnect = _WinHttpConnect($hSession, $node)
	EndIf
EndFunc   ;==>_Steem_Start_Node

; ===========================================================================================
;	_Steem_Close_Node()
;	Close the Connection to the Node
; ===========================================================================================

Func _Steem_Close_Node()
	_WinHttpCloseHandle($hConnect)
	_WinHttpCloseHandle($hSession)
EndFunc   ;==>_Steem_Close_Node

; ===========================================================================================
;	_Steem_Get_Account_Votes($name, $last = 0)
;	$name = The Username without the @
;	$last = the last x Votes 0 = All(Attention can take a long time with accounts that have a lot of votes)
; ===========================================================================================

Func _Steem_Get_Account_Votes($name, $last = 0)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_account_votes","params":["' & $name & '"]}')
	Return _Convert_to_Array(StringTrimLeft(StringTrimRight($Data1, 1), 17), $last)
EndFunc   ;==>_Steem_Get_Account_Votes

; ===========================================================================================
;	_Steem_Get_Reward_Fund_Global()
; ===========================================================================================

Func _Steem_Get_Reward_Fund_Global()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_reward_fund","params":["post"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Reward_Fund_Global

; ===========================================================================================
;	_Steem_Get_Config($hConnect)
; ===========================================================================================

Func _Steem_Get_Config()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_config","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Config

; ===========================================================================================
;	_Steem_Get_State($path)
;	$path = The path
; ===========================================================================================

Func _Steem_Get_State($path)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_state","params":["' & $path & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_State

; ===========================================================================================
;	_Steem_Get_Dynamic_Global_Properties()
; ===========================================================================================

Func _Steem_Get_Dynamic_Global_Properties()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7, "method": "get_dynamic_global_properties", "params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Dynamic_Global_Properties

; ===========================================================================================
;	_Steem_Get_Chain_Properties()
; ===========================================================================================

Func _Steem_Get_Chain_Properties()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7, "method": "get_chain_properties","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Chain_Properties

; ===========================================================================================
;	_Steem_Get_Current_Median_History_Price()
; ===========================================================================================

Func _Steem_Get_Current_Median_History_Price()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_current_median_history_price","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Current_Median_History_Price

; ===========================================================================================
;	_Steem_Get_Feed_History()
; ===========================================================================================

Func _Steem_Get_Feed_History()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_feed_history","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Feed_History

; ===========================================================================================
;	_Steem_Get_Witness_Schedule()
; ===========================================================================================

Func _Steem_Get_Witness_Schedule()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_witness_schedule","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Witness_Schedule

; ===========================================================================================
;	_Steem_Get_Hardfork_Version()
; ===========================================================================================

Func _Steem_Get_Hardfork_Version()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_hardfork_version","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Hardfork_Version

; ===========================================================================================
;	_Steem_Get_Next_Scheduled_Hardfork()
; ===========================================================================================

Func _Steem_Get_Next_Scheduled_Hardfork()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_next_scheduled_hardfork","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Next_Scheduled_Hardfork

; ===========================================================================================
;	_Steem_Get_Account($name)
;	$name = The Account name without the @
; ===========================================================================================

Func _Steem_Get_Account($name)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7, "method": "get_accounts", "params": [["' & $name & '"]]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Account

; ===========================================================================================
;	_Steem_Get_Account_Count()
; ===========================================================================================

Func _Steem_Get_Account_Count()
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_account_count","params":[]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Account_Count

; ===========================================================================================
;	_Steem_Get_Owner_History($name)
;	$name = The Account name without the @
; ===========================================================================================

Func _Steem_Get_Owner_History($name)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_owner_history","params":["' & $name & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Owner_History

; ===========================================================================================
;	_Steem_Get_Block_Header($blockNum)
;	$blockNum = The blocknumber
; ===========================================================================================

Func _Steem_Get_Block_Header($blockNum)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_block_header","params":["' & $blockNum & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Block_Header

; ===========================================================================================
;	_Steem_Get_Block($blocknumber)
;	$blocknumber = The Blocknumber
; ===========================================================================================

Func _Steem_Get_Block($blocknumber)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_block","params":["' & $blocknumber & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Block

; ===========================================================================================
;	_Steem_Get_Vesting_Delegations($account, $from, $limit)
;	$account = The Account name without the @
;	$from = From Account name Without the @ ( For all use "" )
;	$limit = The max Show delegations
; ===========================================================================================

Func _Steem_Get_Vesting_Delegations($account, $from, $limit)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_vesting_delegations","params":["' & $account & '","' & $from & '","' & $limit & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Vesting_Delegations

; ===========================================================================================
;	_Steem_Get_Ops_In_Block($blockNum, $onlyVirtual)
;	$blockNum = The blocknumber
;	$onlyVirtual = True or False
; ===========================================================================================

Func _Steem_Get_Ops_In_Block($blockNum, $onlyVirtual)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_ops_in_block","params":["' & $blockNum & '","' & $onlyVirtual & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Ops_In_Block

; ===========================================================================================
;	_Steem_Get_Content_Replies($name, $permalink)
;	$name = The Account name without the @
;	$permalink = The permalink to the Post
; ===========================================================================================

Func _Steem_Get_Content_Replies($name, $permalink)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_content_replies","params":["' & $name & '", "' & $permalink & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Content_Replies

; ===========================================================================================
;	_Steem_Get_Active_Votes($name, $permalink)
;	$name = The Account name without the @
;	$permalink = The permalink to the Post
; ===========================================================================================

Func _Steem_Get_Active_Votes($name, $permalink)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_active_votes","params":[ "' & $name & '", "' & $permalink & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Active_Votes

; ===========================================================================================
;	_Steem_Get_Content($name, $permalink)
;	$name = The Account name without the @
;	$permalink = The permalink to the Post
; ===========================================================================================

Func _Steem_Get_Content($name, $permalink)
	Local $Data1 = _WinHttpSimpleSSLRequest($hConnect, "POST", "", "", '{"jsonrpc": "2.0","id":7,"method":"get_content","params":[ "' & $name & '", "' & $permalink & '"]}')
	Return _JSONDecode(StringTrimLeft(StringTrimRight($Data1, 1), 17))
EndFunc   ;==>_Steem_Get_Content

; ===========================================================================================
; =================================INTERNAL FUNCTION=========================================
; ===========================================================================================

Func _WinHttpSimpleReadData($hRequest, $iMode = Default)
	__WinHttpDefault($iMode, 0)
	If $iMode > 2 Or $iMode < 0 Then Return SetError(1, 0, '')
	Local $vData = ''
	If $iMode = 2 Then $vData = Binary('')
	If _WinHttpQueryDataAvailable($hRequest) Then
		If $iMode = 0 Then
			Do
				$vData &= _WinHttpReadData($hRequest, 0)
			Until @error
			Return $vData
		Else
			$vData = Binary('')
			Do
				$vData &= _WinHttpReadData($hRequest, 2)
			Until @error
			If $iMode = 1 Then Return BinaryToString($vData, 4)
			Return $vData
		EndIf
	EndIf
	Return SetError(2, 0, $vData)
EndFunc   ;==>_WinHttpSimpleReadData

Func _WinHttpSimpleSSLRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sData = Default, $sHeader = Default, $fGetHeaders = Default, $iMode = Default)
	; Author: ProgAndy
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sData, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	__WinHttpDefault($fGetHeaders, False)
	__WinHttpDefault($iMode, 0)
	If $iMode > 2 Or $iMode < 0 Then Return SetError(4, 0, 0)
	Local $hRequest = _WinHttpSimpleSendSSLRequest($hConnect, $sType, $sPath, $sReferrer, $sData, $sHeader)
	If @error Then Return SetError(@error, 0, 0)
	If $fGetHeaders Then
		Local $aData[2] = [_WinHttpQueryHeaders($hRequest), _WinHttpSimpleReadData($hRequest, $iMode)]
		_WinHttpCloseHandle($hRequest)
		Return $aData
	EndIf
	Local $sOutData = _WinHttpSimpleReadData($hRequest, $iMode)
	_WinHttpCloseHandle($hRequest)
	Return $sOutData
EndFunc   ;==>_WinHttpSimpleSSLRequest

Func __WinHttpDefault(ByRef $vInput, $vOutput)
	If $vInput = Default Or Number($vInput) = -1 Then $vInput = $vOutput
EndFunc   ;==>__WinHttpDefault

Func _WinHttpQueryHeaders($hRequest, $iInfoLevel = Default, $sName = Default, $iIndex = Default)
	__WinHttpDefault($iInfoLevel, $WINHTTP_QUERY_RAW_HEADERS_CRLF)
	__WinHttpDefault($sName, $WINHTTP_HEADER_NAME_BY_INDEX)
	__WinHttpDefault($iIndex, $WINHTTP_NO_HEADER_INDEX)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryHeaders", _
			"handle", $hRequest, _
			"dword", $iInfoLevel, _
			"wstr", $sName, _
			"wstr", "", _
			"dword*", 65536, _
			"dword*", $iIndex)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
	Return SetExtended($aCall[6], $aCall[4])
EndFunc   ;==>_WinHttpQueryHeaders

Func _WinHttpSimpleSendSSLRequest($hConnect, $sType = Default, $sPath = Default, $sReferrer = Default, $sData = Default, $sHeader = Default)
	__WinHttpDefault($sType, "GET")
	__WinHttpDefault($sPath, "")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($sData, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($sHeader, $WINHTTP_NO_ADDITIONAL_HEADERS)
	Local $hRequest = _WinHttpOpenRequest($hConnect, $sType, $sPath, Default, $sReferrer, Default, BitOR($WINHTTP_FLAG_SECURE, $WINHTTP_FLAG_ESCAPE_DISABLE))
	If Not $hRequest Then Return SetError(1, @error, 0)
	If $sType = "POST" And $sHeader = $WINHTTP_NO_ADDITIONAL_HEADERS Then $sHeader = "Content-Type: application/x-www-form-urlencoded" & @CRLF
	_WinHttpSendRequest($hRequest, $sHeader, $sData)
	If @error Then Return SetError(2, 0 * _WinHttpCloseHandle($hRequest), 0)
	_WinHttpReceiveResponse($hRequest)
	If @error Then Return SetError(3, 0 * _WinHttpCloseHandle($hRequest), 0)
	Return $hRequest
EndFunc   ;==>_WinHttpSimpleSendSSLRequest

Func _WinHttpOpenRequest($hConnect, $sVerb = Default, $sObjectName = Default, $sVersion = Default, $sReferrer = Default, $sAcceptTypes = Default, $iFlags = Default)
	__WinHttpDefault($sVerb, "GET")
	__WinHttpDefault($sObjectName, "")
	__WinHttpDefault($sVersion, "HTTP/1.1")
	__WinHttpDefault($sReferrer, $WINHTTP_NO_REFERER)
	__WinHttpDefault($iFlags, $WINHTTP_FLAG_ESCAPE_DISABLE)
	Local $pAcceptTypes
	If $sAcceptTypes = Default Or Number($sAcceptTypes) = -1 Then
		$pAcceptTypes = $WINHTTP_DEFAULT_ACCEPT_TYPES
	Else
		Local $aTypes = StringSplit($sAcceptTypes, ",", 2)
		Local $tAcceptTypes = DllStructCreate("ptr[" & UBound($aTypes) + 1 & "]")
		Local $tType[UBound($aTypes)]
		For $i = 0 To UBound($aTypes) - 1
			$tType[$i] = DllStructCreate("wchar[" & StringLen($aTypes[$i]) + 1 & "]")
			DllStructSetData($tType[$i], 1, $aTypes[$i])
			DllStructSetData($tAcceptTypes, 1, DllStructGetPtr($tType[$i]), $i + 1)
		Next
		$pAcceptTypes = DllStructGetPtr($tAcceptTypes)
	EndIf
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpenRequest", _
			"handle", $hConnect, _
			"wstr", StringUpper($sVerb), _
			"wstr", $sObjectName, _
			"wstr", StringUpper($sVersion), _
			"wstr", $sReferrer, _
			"ptr", $pAcceptTypes, _
			"dword", $iFlags)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_WinHttpOpenRequest

Func _WinHttpSendRequest($hRequest, $sHeaders = Default, $sOptional = Default, $iTotalLength = Default, $iContext = Default)
	__WinHttpDefault($sHeaders, $WINHTTP_NO_ADDITIONAL_HEADERS)
	__WinHttpDefault($sOptional, $WINHTTP_NO_REQUEST_DATA)
	__WinHttpDefault($iTotalLength, 0)
	__WinHttpDefault($iContext, 0)
	Local $pOptional = 0, $iOptionalLength = 0
	If @NumParams > 2 Then
		Local $tOptional
		$iOptionalLength = BinaryLen($sOptional)
		$tOptional = DllStructCreate("byte[" & $iOptionalLength & "]")
		If $iOptionalLength Then $pOptional = DllStructGetPtr($tOptional)
		DllStructSetData($tOptional, 1, $sOptional)
	EndIf
	If Not $iTotalLength Or $iTotalLength < $iOptionalLength Then $iTotalLength += $iOptionalLength
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpSendRequest", _
			"handle", $hRequest, _
			"wstr", $sHeaders, _
			"dword", 0, _
			"ptr", $pOptional, _
			"dword", $iOptionalLength, _
			"dword", $iTotalLength, _
			"dword_ptr", $iContext)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>_WinHttpSendRequest

Func _WinHttpQueryDataAvailable($hRequest)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpQueryDataAvailable", "handle", $hRequest, "dword*", 0)
	If @error Then Return SetError(1, 0, 0)
	Return SetExtended($aCall[2], $aCall[0])
EndFunc   ;==>_WinHttpQueryDataAvailable

Func _WinHttpReadData($hRequest, $iMode = Default, $iNumberOfBytesToRead = Default, $pBuffer = Default)
	__WinHttpDefault($iMode, 0)
	__WinHttpDefault($iNumberOfBytesToRead, 8192)
	Local $tBuffer
	Switch $iMode
		Case 1, 2
			If $pBuffer And $pBuffer <> Default Then
				$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]", $pBuffer)
			Else
				$tBuffer = DllStructCreate("byte[" & $iNumberOfBytesToRead & "]")
			EndIf
		Case Else
			$iMode = 0
			If $pBuffer And $pBuffer <> Default Then
				$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]", $pBuffer)
			Else
				$tBuffer = DllStructCreate("char[" & $iNumberOfBytesToRead & "]")
			EndIf
	EndSwitch
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReadData", _
			"handle", $hRequest, _
			"ptr", DllStructGetPtr($tBuffer), _
			"dword", $iNumberOfBytesToRead, _
			"dword*", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, "")
	If Not $aCall[4] Then Return SetError(-1, 0, "")
	If $aCall[4] < $iNumberOfBytesToRead Then
		Switch $iMode
			Case 0
				Return SetExtended($aCall[4], StringLeft(DllStructGetData($tBuffer, 1), $aCall[4]))
			Case 1
				Return SetExtended($aCall[4], BinaryToString(BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]), 4))
			Case 2
				Return SetExtended($aCall[4], BinaryMid(DllStructGetData($tBuffer, 1), 1, $aCall[4]))
		EndSwitch
	Else
		Switch $iMode
			Case 0, 2
				Return SetExtended($aCall[4], DllStructGetData($tBuffer, 1))
			Case 1
				Return SetExtended($aCall[4], BinaryToString(DllStructGetData($tBuffer, 1), 4))
		EndSwitch
	EndIf
EndFunc   ;==>_WinHttpReadData

Func _WinHttpReceiveResponse($hRequest)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpReceiveResponse", "handle", $hRequest, "ptr", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>_WinHttpReceiveResponse

Func _WinHttpOpen($sUserAgent = Default, $iAccessType = Default, $sProxyName = Default, $sProxyBypass = Default, $iFlag = Default)
	__WinHttpDefault($sUserAgent, "AutoIt/3.3")
	__WinHttpDefault($iAccessType, $WINHTTP_ACCESS_TYPE_NO_PROXY)
	__WinHttpDefault($sProxyName, $WINHTTP_NO_PROXY_NAME)
	__WinHttpDefault($sProxyBypass, $WINHTTP_NO_PROXY_BYPASS)
	__WinHttpDefault($iFlag, 0)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpOpen", _
			"wstr", $sUserAgent, _
			"dword", $iAccessType, _
			"wstr", $sProxyName, _
			"wstr", $sProxyBypass, _
			"dword", $iFlag)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_WinHttpOpen

Func _WinHttpCloseHandle($hInternet)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "bool", "WinHttpCloseHandle", "handle", $hInternet)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return 1
EndFunc   ;==>_WinHttpCloseHandle

Func _WinHttpConnect($hSession, $sServerName, $iServerPort = Default)
	__WinHttpDefault($iServerPort, $INTERNET_DEFAULT_PORT)
	Local $aCall = DllCall($hWINHTTPDLL__WINHTTP, "handle", "WinHttpConnect", _
			"handle", $hSession, _
			"wstr", $sServerName, _
			"dword", $iServerPort, _
			"dword", 0)
	If @error Or Not $aCall[0] Then Return SetError(1, 0, 0)
	Return $aCall[0]
EndFunc   ;==>_WinHttpConnect

Func _Convert_to_Array($string, $last)
	Dim $yourAr
	Local $betw = _StringBetween($string, "{", "}")
	$c = UBound($betw) - 1
	If $last = 0 Then
		For $x = 0 To $c
			$aArray = StringSplit($betw[$x], ",")
			If IsArray($yourAr) = 1 Then
				$Bound = UBound($yourAr)
				ReDim $yourAr[$Bound + 1][5]
				$yourAr[$Bound][0] = StringReplace(StringTrimRight(StringTrimLeft($aArray[1], StringLen('"authorperm":"')), 1), '"', '')
				$yourAr[$Bound][1] = StringReplace(StringTrimLeft($aArray[2], StringLen('"weight":')), '"', '')
				$yourAr[$Bound][2] = StringReplace(StringTrimLeft($aArray[3], StringLen('"rshares":"')), '"', '')
				$yourAr[$Bound][3] = StringReplace(StringTrimLeft($aArray[4], StringLen('"percent":')), '"', '')
				$yourAr[$Bound][4] = StringReplace(StringTrimRight(StringTrimLeft($aArray[5], StringLen('"time":"')), 1), '"', '')
			Else
				Dim $yourAr[2][5]
				$yourAr[0][0] = "authorperm"
				$yourAr[0][1] = "weight"
				$yourAr[0][2] = "rshares"
				$yourAr[0][3] = "percent"
				$yourAr[0][4] = "time"
				$yourAr[1][0] = StringReplace(StringTrimRight(StringTrimLeft($aArray[1], StringLen('"authorperm":"')), 1), '"', '')
				$yourAr[1][1] = StringReplace(StringTrimLeft($aArray[2], StringLen('"weight":')), '"', '')
				$yourAr[1][2] = StringReplace(StringTrimLeft($aArray[3], StringLen('"rshares":"')), '"', '')
				$yourAr[1][3] = StringReplace(StringTrimLeft($aArray[4], StringLen('"percent":')), '"', '')
				$yourAr[1][4] = StringReplace(StringTrimRight(StringTrimLeft($aArray[5], StringLen('"time":"')), 1), '"', '')
			EndIf
		Next
	Else
		If $c - $last <= 0 Then
			$start = 0
		Else
			$start = $c - $last
		EndIf
		For $x = $start To $c
			$aArray = StringSplit($betw[$x], ",")
			If IsArray($yourAr) = 1 Then
				$Bound = UBound($yourAr)
				ReDim $yourAr[$Bound + 1][5]
				$yourAr[$Bound][0] = StringReplace(StringTrimRight(StringTrimLeft($aArray[1], StringLen('"authorperm":"')), 1), '"', '')
				$yourAr[$Bound][1] = StringReplace(StringTrimLeft($aArray[2], StringLen('"weight":')), '"', '')
				$yourAr[$Bound][2] = StringReplace(StringTrimLeft($aArray[3], StringLen('"rshares":"')), '"', '')
				$yourAr[$Bound][3] = StringReplace(StringTrimLeft($aArray[4], StringLen('"percent":')), '"', '')
				$yourAr[$Bound][4] = StringReplace(StringTrimRight(StringTrimLeft($aArray[5], StringLen('"time":"')), 1), '"', '')
			Else
				Dim $yourAr[2][5]
				$yourAr[0][0] = "authorperm"
				$yourAr[0][1] = "weight"
				$yourAr[0][2] = "rshares"
				$yourAr[0][3] = "percent"
				$yourAr[0][4] = "time"
				$yourAr[1][0] = StringReplace(StringTrimRight(StringTrimLeft($aArray[1], StringLen('"authorperm":"')), 1), '"', '')
				$yourAr[1][1] = StringReplace(StringTrimLeft($aArray[2], StringLen('"weight":')), '"', '')
				$yourAr[1][2] = StringReplace(StringTrimLeft($aArray[3], StringLen('"rshares":"')), '"', '')
				$yourAr[1][3] = StringReplace(StringTrimLeft($aArray[4], StringLen('"percent":')), '"', '')
				$yourAr[1][4] = StringReplace(StringTrimRight(StringTrimLeft($aArray[5], StringLen('"time":"')), 1), '"', '')
			EndIf
		Next
	EndIf
	Return $yourAr
EndFunc   ;==>_Convert_to_Array

Func _JSONIsNull($v)
	Return $v == Default
EndFunc   ;==>_JSONIsNull

Func _JSONArray($p0 = 0, $p1 = 0, $p2 = 0, $p3 = 0, $p4 = 0, $p5 = 0, $p6 = 0, $p7 = 0, $p8 = 0, $p9 = 0, $p10 = 0, $p11 = 0, $p12 = 0, $p13 = 0, $p14 = 0, $p15 = 0, $p16 = 0, $p17 = 0, $p18 = 0, $p19 = 0, $p20 = 0, $p21 = 0, $p22 = 0, $p23 = 0, $p24 = 0, $p25 = 0, $p26 = 0, $p27 = 0, $p28 = 0, $p29 = 0, $p30 = 0, $p31 = 0)
	If @NumParams Then
		Local $a[32] = [$p0, $p1, $p2, $p3, $p4, $p5, $p6, $p7, $p8, $p9, $p10, $p11, $p12, $p13, $p14, $p15, $p16, $p17, $p18, $p19, $p20, $p21, $p22, $p23, $p24, $p25, $p26, $p27, $p28, $p29, $p30, $p31]
		ReDim $a[@NumParams]
		Return $a
	EndIf
	Local $d = ObjCreate('Scripting.Dictionary')
	Return $d.keys()
EndFunc   ;==>_JSONArray

Func _JSONIsArray($v)
	Return IsArray($v) And UBound($v, 0) == 1
EndFunc   ;==>_JSONIsArray

Func _JSONObject($p0 = 0, $p1 = 0, $p2 = 0, $p3 = 0, $p4 = 0, $p5 = 0, $p6 = 0, $p7 = 0, $p8 = 0, $p9 = 0, $p10 = 0, $p11 = 0, $p12 = 0, $p13 = 0, $p14 = 0, $p15 = 0, $p16 = 0, $p17 = 0, $p18 = 0, $p19 = 0, $p20 = 0, $p21 = 0, $p22 = 0, $p23 = 0, $p24 = 0, $p25 = 0, $p26 = 0, $p27 = 0, $p28 = 0, $p29 = 0, $p30 = 0, $p31 = 0)
	If @NumParams Then
		Local $a[32] = [$p0, $p1, $p2, $p3, $p4, $p5, $p6, $p7, $p8, $p9, $p10, $p11, $p12, $p13, $p14, $p15, $p16, $p17, $p18, $p19, $p20, $p21, $p22, $p23, $p24, $p25, $p26, $p27, $p28, $p29, $p30, $p31]
		ReDim $a[@NumParams]
		Return _JSONObjectFromArray($a)
	EndIf
	Return _JSONObjectFromArray(0)
EndFunc   ;==>_JSONObject

Func _JSONObjectFromArray($a)
	Local $o[1][2] = [[$_JSONNull, 'JSONObject']], $len = UBound($a)
	If $len Then
		ReDim $o[Floor($len / 2) + 1][2]
		Local $oi = 1
		Local $d = ObjCreate('Scripting.Dictionary')
		For $ai = 1 To $len - 1 Step 2
			Local $k = String($a[$ai - 1])
			If $d.exists($k) Then
				Return SetError(1, $d.count + 1, 0)
			EndIf
			$d.add($k, True)
			$o[$oi][0] = $k
			$o[$oi][1] = $a[$ai]
			$oi += 1
		Next
	EndIf
	Return $o
EndFunc   ;==>_JSONObjectFromArray

Func _JSONIsObject($v)
	If IsArray($v) And UBound($v, 0) == 2 And UBound($v, 2) == 2 Then
		Return _JSONIsNull($v[0][0]) And $v[0][1] == 'JSONObject'
	EndIf
	Return False
EndFunc   ;==>_JSONIsObject

Global $_JSONErrorMessage = ''
Local $__JSONTranslator
Local $__JSONReadNextFunc, $__JSONOffset, $__JSONAllowDuplicatedObjectKeys
Local $__JSONCurr, $__JSONWhitespaceWasFound
Local $__JSONDecodeString, $__JSONDecodePos
Local $__JSONIndentString, $__JSONIndentLen, $__JSONComma, $__JSONColon
Local $__JSONEncodeErrFlags, $__JSONEncodeErrCount

Func _JSONDecodeWithReadFunc($funcName, $translator = '', $allowDuplicatedObjectKeys = False, $postCheck = False)
	$_JSONErrorMessage = ''
	If Not __JSONSetTranslator($translator) Then
		Return SetError(999, 0, 0)
	EndIf
	$__JSONReadNextFunc = $funcName
	$__JSONOffset = 0
	$__JSONAllowDuplicatedObjectKeys = $allowDuplicatedObjectKeys
	__JSONReadNext()
	If @error Then
		Return SetError(@error, @extended, 0)
	EndIf
	Local $v = __JSONDecodeInternal()
	If @error Then
		Return SetError(@error, @extended, 0)
	EndIf
	If $postCheck Then
		If __JSONSkipWhitespace() Then
			$_JSONErrorMessage = 'string contains unexpected text after the decoded JSON value'
			Return SetError(99, 0, 0)
		EndIf
	EndIf
	If $translator Then
		$v = __JSONDecodeTranslateWalk($_JSONNull, $_JSONNull, $v)
	EndIf
	If @error Then
		Return SetError(@error, @extended, $v)
	EndIf
	Return $v
EndFunc   ;==>_JSONDecodeWithReadFunc

Func _JSONDecode($s, $translator = '', $allowDuplicatedObjectKeys = False, $startPos = 1)
	$__JSONDecodeString = String($s)
	$__JSONDecodePos = $startPos
	Local $v = _JSONDecodeWithReadFunc('__JSONReadNextFromString', $translator, $allowDuplicatedObjectKeys, True)
	If @error Then
		Return SetError(@error, @extended, $v)
	EndIf
	Return $v
EndFunc   ;==>_JSONDecode

Func _JSONDecodeAll($s, $translator = '', $allowDuplicatedObjectKeys = False)
	Local $v = _JSONDecode('[' & $s & @LF & ']', $translator, $allowDuplicatedObjectKeys)
	If @error Then
		Return SetError(@error, @extended, $v)
	EndIf
	Return $v
EndFunc   ;==>_JSONDecodeAll

Func _JSONEncode($v, $translator = '', $indent = '', $linebreak = @CRLF, $omitCommas = False)
	$_JSONErrorMessage = ''
	If Not __JSONSetTranslator($translator) Then
		Return SetError(999, 0, 0)
	EndIf
	If $indent And $linebreak Then
		If IsBool($indent) Then
			$__JSONIndentString = @TAB
		Else
			$__JSONIndentString = String($indent)
		EndIf
		$__JSONIndentLen = StringLen($__JSONIndentString)
		$__JSONColon = ': '
		If $omitCommas Then
			$__JSONComma = ''
		Else
			$__JSONComma = ','
		EndIf
	Else
		$indent = ''
		$linebreak = ''
		$__JSONColon = ':'
		$__JSONComma = ','
	EndIf
	$__JSONEncodeErrFlags = 0
	$__JSONEncodeErrCount = 0
	Local $s = __JSONEncodeInternal($_JSONNull, $_JSONNull, $v, $linebreak) & $linebreak ; start indentation with the linebreak character
	If @error Then
		Return SetError(@error, @extended, '')
	EndIf
	If $__JSONEncodeErrCount Then
		Return SetError($__JSONEncodeErrFlags, $__JSONEncodeErrCount, $s)
	EndIf
	Return $s
EndFunc   ;==>_JSONEncode

Func __JSONSetTranslator($translator)
	If $translator Then
		Local $dummy = Call($translator, 0, 0, 0)
		If @error == 0xDEAD And @extended == 0xBEEF Then
			$_JSONErrorMessage = 'translator function not defined, or defined with wrong number of parameters'
			Return False
		EndIf
	EndIf
	$__JSONTranslator = $translator
	Return True
EndFunc   ;==>__JSONSetTranslator

Func __JSONReadNext($numChars = 1)
	$__JSONCurr = Call($__JSONReadNextFunc, $numChars)
	If @error Then
		If @error == 0xDEAD And @extended == 0xBEEF Then
			$_JSONErrorMessage = 'read function not defined, or defined with wrong number of parameters'
		EndIf
		Return SetError(@error, @extended, 0)
	EndIf
	$__JSONOffset += StringLen($__JSONCurr)
	Return $__JSONCurr
EndFunc   ;==>__JSONReadNext

Func __JSONReadNextFromString($numChars)
	Local $s = StringMid($__JSONDecodeString, $__JSONDecodePos, $numChars)
	$__JSONDecodePos += $numChars
	Return $s
EndFunc   ;==>__JSONReadNextFromString

Func __JSONSkipWhitespace()
	$__JSONWhitespaceWasFound = False
	While $__JSONCurr
		If StringIsSpace($__JSONCurr) Then
			$__JSONWhitespaceWasFound = True
		ElseIf $__JSONCurr == '/' Then
			Switch __JSONReadNext()
				Case '/'
					While __JSONReadNext() And Not StringRegExp($__JSONCurr, "[\n\r]", 0)
					WEnd
					If $__JSONCurr Then
						$__JSONWhitespaceWasFound = True
					Else
						Return ''
					EndIf
				Case '*'
					__JSONReadNext()
					While $__JSONCurr
						If $__JSONCurr == '*' Then
							If __JSONReadNext() == '/' Then
								ExitLoop
							EndIf
						Else
							__JSONReadNext()
						EndIf
					WEnd
					If Not $__JSONCurr Then
						$_JSONErrorMessage = 'unterminated block comment'
						ExitLoop
					EndIf
				Case Else
					$_JSONErrorMessage = 'bad comment syntax'
					ExitLoop
			EndSwitch
		Else
			Return $__JSONCurr
		EndIf
		__JSONReadNext()
	WEnd
	Return SetError(2, 0, 0)
EndFunc   ;==>__JSONSkipWhitespace

Func __JSONDecodeObject()
	Local $d = ObjCreate('Scripting.Dictionary'), $key
	Local $o = _JSONObject(), $len = 1, $i, $k
	If $__JSONCurr == '{' Then
		__JSONReadNext()
		If __JSONSkipWhitespace() == '}' Then
			__JSONReadNext()
			Return $o
		EndIf
		While $__JSONCurr
			$key = __JSONDecodeObjectKey()
			If @error Then
				Return SetError(@error, @extended, 0)
			EndIf
			If __JSONSkipWhitespace() <> ':' Then
				$_JSONErrorMessage = 'expected ":", encountered "' & $__JSONCurr & '"'
				ExitLoop
			EndIf
			If $d.exists($key) Then
				If $__JSONAllowDuplicatedObjectKeys Then
					$i = $d.item($key)
				Else
					$_JSONErrorMessage = 'duplicate key specified for object: "' & $k & '"'
					ExitLoop
				EndIf
			Else
				$i = $len
				$len += 1
				ReDim $o[$len][2]
				$o[$i][0] = $key
				$d.add($key, $i)
			EndIf
			__JSONReadNext()
			$o[$i][1] = __JSONDecodeInternal()
			If @error Then
				Return SetError(@error, @extended, 0)
			EndIf
			Switch __JSONSkipWhitespace()
				Case '}'
					__JSONReadNext()
					Return $o
				Case ','
					__JSONReadNext()
					__JSONSkipWhitespace()
				Case Else
					If Not $__JSONWhitespaceWasFound Then
						$_JSONErrorMessage = 'expected "," or "}", encountered "' & $__JSONCurr & '"'
						ExitLoop
					EndIf
			EndSwitch
		WEnd
	EndIf
	Return SetError(3, 0, 0)
EndFunc   ;==>__JSONDecodeObject

Func __JSONDecodeObjectKey()
	If $__JSONCurr == '"' Or $__JSONCurr == "'" Then
		Return __JSONDecodeString()
	EndIf
	If StringIsDigit($__JSONCurr) Then
		Return String(__JSONDecodeNumber())
	EndIf
	Local $s = ''
	While (StringIsAlNum($__JSONCurr) Or $__JSONCurr == '_')
		$s &= $__JSONCurr
		__JSONReadNext()
	WEnd
	If Not $s Then
		$_JSONErrorMessage = 'expected object key, encountered "' & $__JSONCurr & '"'
		Return SetError(13, 0, 0)
	EndIf
	Return $s
EndFunc   ;==>__JSONDecodeObjectKey

Func __JSONDecodeArray()
	Local $a = _JSONArray(), $len = 0
	If $__JSONCurr == '[' Then
		__JSONReadNext()
		If __JSONSkipWhitespace() == ']' Then
			__JSONReadNext()
			Return $a
		EndIf
		While $__JSONCurr
			$len += 1
			ReDim $a[$len]
			$a[$len - 1] = __JSONDecodeInternal()
			If @error Then
				Return SetError(@error, @extended, 0)
			EndIf
			Switch __JSONSkipWhitespace()
				Case ']'
					__JSONReadNext()
					Return $a
				Case ','
					__JSONReadNext()
					__JSONSkipWhitespace()
				Case Else
					If Not $__JSONWhitespaceWasFound Then
						$_JSONErrorMessage = 'expected "," or "]", encountered "' & $__JSONCurr & '"'
						ExitLoop
					EndIf
			EndSwitch
		WEnd
	EndIf
	Return SetError(4, 0, 0)
EndFunc   ;==>__JSONDecodeArray

Func __JSONDecodeString()
	Local $s = '', $q = $__JSONCurr
	If $q == '"' Or $q == "'" Then
		While $__JSONCurr
			__JSONReadNext()
			Select
				Case $__JSONCurr == $q
					__JSONReadNext()
					Return $s
				Case $__JSONCurr == '\'
					Switch __JSONReadNext()
						Case '\', '/', '"', "'"
							$s &= $__JSONCurr
						Case 't'
							$s &= @TAB
						Case 'n'
							$s &= @LF
						Case 'r'
							$s &= @CR
						Case 'f'
							$s &= ChrW(0xC)
						Case 'b'
							$s &= ChrW(0x8)
						Case 'v'
							$s &= ChrW(0xB)
						Case 'u'
							If StringIsXDigit(__JSONReadNext(4)) Then
								$s &= ChrW(Dec($__JSONCurr))
							Else
								ExitLoop
							EndIf
						Case Else
							ExitLoop
					EndSwitch
				Case AscW($__JSONCurr) >= 0x20
					$s &= $__JSONCurr
				Case Else
					ExitLoop
			EndSelect
		WEnd
	EndIf
	Return SetError(5, 0, 0)
EndFunc   ;==>__JSONDecodeString

Func __JSONDecodeHexNumber($negative)
	Local $n = 0
	While StringIsXDigit(__JSONReadNext())
		$n = $n * 0x10 + Dec($__JSONCurr)
	WEnd
	If $negative Then
		Return -$n
	EndIf
	Return $n
EndFunc   ;==>__JSONDecodeHexNumber

Func __JSONDecodeNumber()
	Local $s = ''
	If $__JSONCurr == '+' Or $__JSONCurr == '-' Then
		$s &= $__JSONCurr
		__JSONReadNext()
	EndIf
	If $__JSONCurr == '0' Then
		$s &= $__JSONCurr
		__JSONReadNext()
		If StringLower($__JSONCurr) == 'x' Then
			Return __JSONDecodeHexNumber(StringLeft($s, 1) == '-')
		EndIf
	EndIf
	While StringIsDigit($__JSONCurr)
		$s &= $__JSONCurr
		__JSONReadNext()
	WEnd
	If $__JSONCurr == '.' Then
		$s &= $__JSONCurr
		While StringIsDigit(__JSONReadNext())
			$s &= $__JSONCurr
		WEnd
	EndIf
	If StringLower($__JSONCurr) == 'e' Then
		$s &= $__JSONCurr
		__JSONReadNext()
		If $__JSONCurr == '+' Or $__JSONCurr == '-' Then
			$s &= $__JSONCurr
			__JSONReadNext()
		EndIf
		While StringIsDigit($__JSONCurr)
			$s &= $__JSONCurr
			__JSONReadNext()
		WEnd
		Return Execute($s)
	EndIf
	Return Number($s)
EndFunc   ;==>__JSONDecodeNumber

Func __JSONDecodeLiteral()
	Switch $__JSONCurr
		Case 't'
			If __JSONReadNext(3) == 'rue' Then
				__JSONReadNext()
				Return True
			EndIf
		Case 'f'
			If __JSONReadNext(4) == 'alse' Then
				__JSONReadNext()
				Return False
			EndIf
		Case 'n'
			If __JSONReadNext(3) == 'ull' Then
				__JSONReadNext()
				Return $_JSONNull
			EndIf
	EndSwitch
	Return SetError(7, 0, 0)
EndFunc   ;==>__JSONDecodeLiteral

Func __JSONDecodeInternal()
	Local $v
	Switch __JSONSkipWhitespace()
		Case '{'
			$v = __JSONDecodeObject()
		Case '['
			$v = __JSONDecodeArray()
		Case '"', "'"
			$v = __JSONDecodeString()
		Case '0' To '9', '-', '+'
			$v = __JSONDecodeNumber()
		Case Else
			$v = __JSONDecodeLiteral()
	EndSwitch
	If @error Then
		Return SetError(@error, @extended, $v)
	EndIf
	Return $v
EndFunc   ;==>__JSONDecodeInternal

Func __JSONDecodeTranslateWalk(Const ByRef $holder, Const ByRef $key, $value)
	Local $v
	If IsArray($value) Then
		If _JSONIsObject($value) Then
			For $i = 1 To UBound($value) - 1
				$v = __JSONDecodeTranslateWalk($value, $value[$i][0], $value[$i][1])
				Switch @error
					Case 0
						$value[$i][1] = $v
					Case 4627
						$value[$i][0] = $_JSONNull
					Case Else
						Return SetError(@error, @extended, 0)
				EndSwitch
			Next
		Else
			For $i = 0 To UBound($value) - 1
				$v = __JSONDecodeTranslateWalk($value, $i, $value[$i])
				Switch @error
					Case 0
						$value[$i] = $v
					Case 4627
						$value[$i] = $_JSONNull
					Case Else
						Return SetError(@error, @extended, 0)
				EndSwitch
			Next
		EndIf
	EndIf
	$v = Call($__JSONTranslator, $holder, $key, $value)
	If @error Then
		Return SetError(@error, @extended, 0)
	EndIf
	Return $v
EndFunc   ;==>__JSONDecodeTranslateWalk

Func __JSONEncodeObject($o, Const ByRef $indent)
	Local $result = '', $inBetween = $__JSONComma & $indent, $d = ObjCreate('Scripting.Dictionary')
	For $i = 1 To UBound($o) - 1
		Local $key = $o[$i][0]
		If Not _JSONIsNull($key) Then
			$key = String($key)
			If $d.exists($key) Then
				$__JSONEncodeErrFlags = BitOR($__JSONEncodeErrFlags, 2)
				$__JSONEncodeErrCount += 1
			Else
				$d.add($key, True)
				Local $s = __JSONEncodeInternal($o, $key, $o[$i][1], $indent)
				If @error Then
					Return SetError(@error, @extended, $result)
				EndIf
				If $s Then
					$result &= $inBetween & __JSONEncodeString($key) & $__JSONColon & $s
				EndIf
			EndIf
		EndIf
	Next
	If $indent And $result Then
		$result &= StringTrimRight($indent, $__JSONIndentLen)
	EndIf
	Return '{' & StringTrimLeft($result, StringLen($__JSONComma)) & '}'
EndFunc   ;==>__JSONEncodeObject

Func __JSONEncodeArray($a, Const ByRef $indent)
	Local $result = '', $inBetween = $__JSONComma & $indent
	If UBound($a) Then
		For $i = 0 To UBound($a) - 1
			Local $s = __JSONEncodeInternal($a, $i, $a[$i], $indent)
			If @error Then
				Return SetError(@error, @extended, $result)
			EndIf
			If Not $s Then
				$s = 'null'
			EndIf
			$result &= $inBetween & $s
		Next
		If $indent Then
			$result &= StringTrimRight($indent, $__JSONIndentLen)
		EndIf
	EndIf
	Return '[' & StringTrimLeft($result, StringLen($__JSONComma)) & ']'
EndFunc   ;==>__JSONEncodeArray

Func __JSONEncodeString($s)
	Local Const $escape = '[\x00-\x1f\x7f-\x9f\u00ad\u0600-\u0604\u070f\u17b4\u17b5\u200c-\u200f\u2028-\u202f\u2060-\u206f\ufeff\ufff0-\uffff]'
	Local $result, $ch, $u
	$s = StringRegExpReplace($s, '([\"\\])', '\\\0')
	If StringRegExp($s, $escape, 0) Then
		$result = ''
		For $i = 1 To StringLen($s)
			$ch = StringMid($s, $i, 1)
			If StringRegExp($ch, $escape, 0) Then
				$u = AscW($ch)
				Switch $u
					Case 0x9
						$ch = '\t'
					Case 0xA
						$ch = '\n'
					Case 0xD
						$ch = '\r'
					Case 0xC
						$ch = '\f'
					Case 0x8
						$ch = '\b'
					Case Else
						$ch = '\u' & Hex($u, 4)
				EndSwitch
			EndIf
			$result &= $ch
		Next
	Else
		$result = $s
	EndIf
	Return '"' & $result & '"'
EndFunc   ;==>__JSONEncodeString

Func __JSONEncodeInternal(Const ByRef $holder, Const ByRef $k, $v, $indent)
	Local $s
	If $indent Then
		$indent &= $__JSONIndentString
		If StringLen($indent) / $__JSONIndentLen > 255 Then
			$_JSONErrorMessage = 'max depth exceeded – possible data recursion'
			Return SetError(1, 0, 0)
		EndIf
	EndIf
	If $__JSONTranslator Then
		$v = Call($__JSONTranslator, $holder, $k, $v)
		Switch @error
			Case 0
			Case 4627
				Return ''
			Case Else
				Return SetError(@error, @extended, '')
		EndSwitch
	EndIf
	Select
		Case _JSONIsObject($v)
			$s = __JSONEncodeObject($v, $indent)
		Case _JSONIsArray($v)
			$s = __JSONEncodeArray($v, $indent)
		Case IsString($v)
			$s = __JSONEncodeString($v)
		Case IsNumber($v)
			$s = String($v)
		Case IsBool($v)
			$s = StringLower($v)
		Case _JSONIsNull($v)
			$s = 'null'
		Case Else
			$__JSONEncodeErrFlags = BitOR($__JSONEncodeErrFlags, 1)
			$__JSONEncodeErrCount += 1
			$s = 'null'
	EndSelect
	If @error Then
		Return SetError(@error, @extended, $s)
	EndIf
	Return $s
EndFunc   ;==>__JSONEncodeInternal

