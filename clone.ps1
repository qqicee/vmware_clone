#命令语法检测
param(
[string]$filename=$(throw "参数丢失Parameter Missing: -filename target_excel"),
[string]$vchost=$(throw "参数丢失Parameter Missing: -vchost ip_address")
)

# 配置																											#
#####################################################################################################################
$WorkPool =  3
$diff_size = 50    # 当模板硬盘和克隆的新虚拟机硬盘大小，相差多少才添加硬盘，防止添加小硬盘。小于该值，原盘增加，大于该值，添加新的。
$vcenterhost = $vchost
$spec="temp_spec"
$doman="localdomain"
$FAILED=0     #用于记录失败次数
$now= get-date -Format 'yyyy-MM-dd HH:mm:ss'

# 目标文件检测  
#####################################################################################################################
$filename= Split-Path -leaf $filename
$FilePath= resolve-path $filename  2>$null
if ($filepath -eq $null )  {
	write-host $now 文件: $Filename 不存在，退出 -foregroundcolor red
	exit 13	
}
elseif (test-path $filepath)
{
	write-host $now 使用: $FilePath 作为任务列表 -foregroundcolor green
}

# 加载Powercli模块
#####################################################################################################################
if (-not( get-module -name VMware.VimAutomation.Core))	{ 
	Write-Host "$now 正在初始化..........."
	import-module VMware.Powercli
	if ($? -eq "true") {
		write-host "$now 初始化成功" -foregroundcolor green 
		}	
	else
		{
		write-host "$now 初始化失败,检查是否已经安装了合适版本的Powercli,按任意键退出" -foregroundcolor red 
		read-host
		exit 12  # 加载Powercli模块失败
		}
}
else 
{
    write-host "$now 初始化成功"  -foregroundcolor green 
}

# 连接vc																											#
#####################################################################################################################
write-host "$now 正在连接 $vcenterhost " -foregroundcolor green 
$RETAL=Connect-VIServer -Server $vcenterhost #-username $vcenterusr -Password $vcenterpsk
if ( ! $RETAL.IsConnected) {
    write-host "...$now 连接 $vcenterhost 失败 " -foregroundcolor red
    $FAILED++
    exit 11   # 连接vc失败。
    }
else {
        write-host "...$now 连接 $vcenterhost 成功，开始工作。 " -foregroundcolor green
		}

if (! (Get-OSCustomizationSpec $spec 2>$null )){
New-OSCustomizationSpec -name $spec -Type NonPersistent -OSType Linux -Domain $doman -NamingScheme fixed -NamingPrefix localhost 
}
$custsysprep = Get-OSCustomizationSpec $spec

# 读取 Excel																									    #
#####################################################################################################################
Stop-Process -Name et,excel  2>$null   # Wps:et   Excel:excel
$objExcel = New-Object -ComObject Excel.Application
$objExcel.Visible = $false
$WorkBook = $objExcel.Workbooks.Open($FilePath, $true)  
$WorkSheet = $WorkBook.Sheets.Item(1)
$x = 1
do
{
    $x = $x + 1
	# Excel 列1--Esxi 计算资源
    $esxihost = $WorkSheet.cells.item($x,1).Text.trim()

    if ( $esxihost -ne "" ) {
	# Excel 列2--虚拟机名称
		$vmname =  $WorkSheet.cells.item($x,2).Text.trim()
	# Excel 列3--模板
		$source = $WorkSheet.cells.item($x,3).Text.trim()
		if ( $source -ne "" ) {
			if (get-vm $source) {$source_tag= "-vm "} else {$source_tag = "-template "}
			$source= $source_tag + $source 
		}
	# Excel 列4--数据存储
		$datastore = $WorkSheet.cells.item($x,4).Text.trim()
	# Excel 列5--虚拟机主机名
		$hostname = $WorkSheet.cells.item($x,5).Text.trim()
		if ( $hostname -eq "" ) { $hostname="localhost" }
	# Excel 列6、7、8--数据IP、掩码、网关
		$ip = $WorkSheet.cells.item($x,6).Text.trim()
		$mask =     $WorkSheet.cells.item($x,7).Text.trim()
		$gateway =  $WorkSheet.cells.item($x,8).Text.trim()
	# Excel 列9--网络表情
		$net_tag =  $WorkSheet.cells.item($x,9).Text.trim()
	# Excel 列10--虚拟机文件夹
		$location = $WorkSheet.cells.item($x,10).Text.trim()
		# 文件夹不存在就创建
		if ($location -ne "") {
			if (!(get-folder $location -ErrorAction Ignore )) {
				write-host "...$now 文件夹 $location 不存在，开始创建 虚拟机文件夹。 " -foregroundcolor green
				New-Folder -Name $location -Location (Get-Folder vm) 

			}
			$location="-location $location"
		}
		else { 
			$location = ""
		}
        
	# Excel 列11--内存大小-GB
		$memsize =  $WorkSheet.cells.item($x,11).Text.trim()
        if ($memsize -ne "") {$chmem_tag=$true} else {$chmem_tag=$false}
	# Excel 列12--CPU核数-C
		$cpunum =   $WorkSheet.cells.item($x,12).Text.trim()
        if ($cpunum -ne "") {$chcpu_tag=$true} else {$chcpu_tag=$false}
	# Excel 列13--硬盘总大小-GB
		$disksize =  $WorkSheet.cells.item($x,13).Text.trim()
		if ($disksize -ne "") {$chdisk_tag=$true} else {$chdisk_tag=$false}

				
		$custsysprep | Set-OScustomizationSpec -NamingScheme fixed -NamingPrefix $hostname   1>$null
        $custsysprep | Get-OSCustomizationNicMapping | Set-OSCustomizationNicMapping -IpMode UseStaticIP -IpAddress $ip -SubnetMask $mask -DefaultGateway $gateway  1>$null
        Write-Host "`n$now 开始从" ($source -split " ")[1] "---克隆---> $vmname " -ForegroundColor Green
        Invoke-Expression  "New-vm -name $vmname -vmhost $esxihost  $source -datastore $datastore -OSCustomizationspec $custsysprep $location -diskstorageformat thick"  1>$null
	    
        if ( $? -eq "true" )
        {
        Write-Host "---$now 克隆虚拟机 $vmname 成功" -ForegroundColor Green

        #更改 内存，cpu
        if ($chmem_tag -and $chcpu_tag ) {
            Write-Host "---$now 指定了CPU、内存配置，开始为 $vmname 配置CPU、内存"　-ForegroundColor Green
            get-vm $vmname |set-vm -memoryGB $memsize -numcpu $cpunum -Confirm:$False  1>$null
        }
        elseif ($chmem_tag) {
            Write-Host "---$now 指定了内存配置,为 $vmname 内存" -ForegroundColor Green 
            get-vm $vmname |set-vm -memoryGB $memsize  -Confirm:$Fals  1>$null
        }
        elseif ($chcpu_tag) {
            Write-Host "---$now 指定了CPU配置，为 $vmname 配置CPU" -ForegroundColor Green
            get-vm $vmname |set-vm -numcpu $cpunum -Confirm:$False   1>$null
        }
		# 更改硬盘
		if ($chdisk_tag ) {
			Write-Host "---$now 指定了硬盘配置，开始为 $vmname 配置硬盘"　-ForegroundColor Green
			$current_size= (get-vm $vmname|get-harddisk|Measure-Object CapacityGB -sum).sum
			$remain_size=$disksize-$current_size
			if ($remain_size -lt $diff_size )   {   # 原位添加
			    Write-Host "---$now 原硬盘扩展 $remain_size GB"　-ForegroundColor Green
				$last_disk= get-vm $vmname|get-harddisk|select -last 1
				$last_disk_size=($last_disk|select -Property capacityGB).capacityGB
				$last_disk|set-harddisk -capacityGB ($last_disk_size+$remain_size) -Confirm:$False   1>$null
			}  
			else   #> diff_size ,添加新硬盘
			{
		    Write-Host "---$now 新加硬盘 $remain_size GB"　-ForegroundColor Green
			get-vm $vmname |new-harddisk -capacityGB $remain_size -Confirm:$False   1>$null	
			}
		}
		
        Write-Host "---$now 开始为 $vmname 配置网络" -ForegroundColor Green        
        Get-VM $vmname | Get-NetworkAdapter | Set-NetworkAdapter -NetworkName $net_tag -Confirm:$false 1>$null
            if ( $? -eq "true" )
            {
				$WorkSheet.cells.item($x,14) = "Success"
            }
            else
            {
                $WorkSheet.cells.item($x,14) = "net-Failed"
            }
			
            start-vm -vm $vmname 1>$null
			Write-Host "---$now 克隆 $vmname 完成`n" -ForegroundColor Green        
        }
        else
        {
            $WorkSheet.cells.item($x,14) = "clone-Failed"
        }
	$WorkBook.Save() | Out-Null
    }
}
while($esxihost -ne "")
#####################################################################################################################
# Scripts Exit													    #
#####################################################################################################################
$WorkBook.Close()  
$objExcel.Quit()
Stop-Process -Name et,excel  2>$null
write-host "虚拟机克隆全部完成" -foregroundcolor Green