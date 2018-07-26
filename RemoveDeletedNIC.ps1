##################################################################
# 取り外した NIC をデバイスリストから削除する
##################################################################
# スクリプトディレクトリ
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# devcon.exe のプルパス
$DevCon = "C:\Program Files (x86)\Windows Kits\10\Tools\x64\devcon.exe"


##################################################################
# 現在の NIC 取得
##################################################################
function GetNowNICs(){
	# 現在の NIC リスト
	[array]$NowNICs = . "$DevCon" listclass net

	# 現在の NIC リストを格納するオブジェクト
	$NowNICObjects = @()

	# オブジェクトにする
	$Max = $NowNICs.Count
	for( $Index = 0;$Index -lt $Max; $Index++){
		$NowNICObject = New-Object PSObject | Select-Object DeviceID, DeviceName

		if( $Index -eq 0 ){
			$NowNICs[$Index] = $NowNICs[$Index] -replace "^Listing [0-9]+ .+\("
		}

		[array]$Buffer = $NowNICs[$Index] -Split ":"
		if($Buffer.Count -eq 2){
			$NowNICObject.DeviceID = $Buffer[0].Trim()
			$NowNICObject.DeviceName = $Buffer[1].Trim()
		}
		$NowNICObjects += $NowNICObject
	}

	return $NowNICObjects
}

##################################################################
# 未接続も含めた全 NIC 取得
##################################################################
function GetAllNICs(){
	# 未接続も含めた全 NIC リスト
	[array]$AllNICs = . "$DevCon" findall =net

	# 未接続も含めた全 NIC リストを格納するオブジェクト
	$AllNICObjects = @()

	# オブジェクトにする
	$Max = $AllNICs.Count -1
	for( $Index = 0;$Index -lt $Max; $Index++){
		$AllNICObject = New-Object PSObject | Select-Object DeviceID, DeviceName

		[array]$Buffer = $AllNICs[$Index] -Split ":"
		if($Buffer.Count -eq 2){
			$AllNICObject.DeviceID = $Buffer[0].Trim()
			$AllNICObject.DeviceName = $Buffer[1].Trim()
		}
		$AllNICObjects += $AllNICObject
	}

	return $AllNICObjects
}

##################################################################
# 未接続 NIC をデバイスから削除する
##################################################################
function RemoveNotConnectNIC( $NowNICObjects, $AllNICObjects){
	# 未接続 NIC をデバイスから削除する
	$Max = $AllNICObjects.Count
	for( $Index = 0;$Index -lt $Max; $Index++){
		# 未接続になっている NIC
		if(-not $NowNICObjects.DeviceID.Contains($AllNICObjects[$Index].DeviceID)){
			# RAS 以外の未接続 NIC を消す
			if( $AllNICObjects[$Index].DeviceName -ne "RAS"){
				$RemoveNICID = $AllNICObjects[$Index].DeviceID
				$RemoveNICName = $AllNICObjects[$Index].DeviceName
				echo "Removed NIC / $RemoveNICID : $RemoveNICName"

				# devcon.exe で削除
				$RemoveTerget = "`"@" + $RemoveNICID + "`""
				. $DevCon -r remove $RemoveTerget
			}
		}
	}
}

##################################################################
# main
##################################################################

# 現在有効な NIC 取得
[array]$NowNICs = GetNowNICs

# 取り外した分も含めたすべての NIC
[array]$AllNICs = GetAllNICs

# 取り外された NIC 削除
RemoveNotConnectNIC $NowNICs $AllNICs
