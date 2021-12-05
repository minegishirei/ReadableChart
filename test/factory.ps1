
function debug($code){
    (Read-Host "debug : "  $code)
}



class Parts{

    ##param
    $children = [System.Collections.ArrayList]::new()
    [string]$name
    [string]$contents
    [Parts]$mother;

    ##method
    [void]addChild($parts){
    
        $this.children.Add($parts)
    }

}



class FolderParts : Parts{

}


class FileParts : Parts{
    [System.Xml.XmlDocument]AppendXmlChild($container, [System.Xml.XmlDocument]$doc){
        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
            $child_xml = $doc.CreateNode("element", "child", $xmlns)
        $container.AppendChild($child_xml) | Out-Null

        ##子アイテム
        foreach ($item in $this.children) {
            $item.AppendXmlChild($child_xml, $doc)
        }
        return $container
    }

    [void] setName($filename){
        $this.name = $filename
    }

}


class ContextParts : Parts {
    $code_list = [System.Collections.ArrayList]::new()
    [void]getchildren(){
        foreach ($item in $this.children) {
            foreach ($code in $this.code_list) {
                Write-Host $code
            }
            Write-Host "=----------------------------#"
            $item.getchildren()
        }
    }
    [System.Xml.XmlDocument]AppendXmlChild($container, [System.Xml.XmlDocument]$doc){
        #[System.Xml.XmlDocument]$doc = [System.Xml.XmlDocument]::new()
        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
        $child_xml = $doc.CreateNode("element", "child", $xmlns)
            $name_xml = $doc.CreateNode("element", "name", $xmlns)
                $name_xml.InnerText = $this.name
            $child_xml.AppendChild($name_xml)
            $code_xml = $doc.CreateNode("element", "code_list", $xmlns)
                $code = "`n"
                foreach ($item in $this.code_list) {
                    $code += ($item+"`n")
                }
                $code_xml.InnerText = $code
            $child_xml.AppendChild($code_xml) | Out-Null
        $container.AppendChild($child_xml) | Out-Null

        ##子アイテム
        foreach ($item in $this.children) {
            $item.AppendXmlChild($child_xml, $doc)
        }
        return $doc
    }
}


class FunctionParts : ContextParts {
    [string]$type = "function";
    [bool]$is_sub_start = $False;
    [bool]$is_function_start = $False;

    [bool]isstart([string]$line){
        return ( ($line -like "*Sub *") -and ($line -notlike "*End*Sub*" ) )
    }

    [bool]isend([string]$line){
        return $line -like "*End*Sub*"
    }
    
    [void]setName([string]$functionline){
        try{
            $right = $functionline.Split("Sub")[1]
            $middle = $right.Split("(")[0]
            $this.name = $middle.Trim()
        }catch{
            Write-Output "Something threw an exception or used Write-Error"
            Write-Output $_
        }
    }
}




class ContextFactory : ContextParts{

    #method
    [bool]isend($line){
        return $false
    }
    [bool]isstart($line){
        return $True
    }
    [void]setName($name){
        $this.name = $name
    }

    ##mainmethod
    [System.Collections.ArrayList]run($filepath){
        $controller = $this;
        $contents = (Get-Content $filepath)
        foreach ($line in $contents) {
            $functionParts = [FunctionParts]::new();
            if($functionParts.isstart($line)){
                $controller.addChild($functionParts)
                $functionParts.setName($line)
                $functionParts.mother = $controller
                $controller = $functionParts
            }
            $controller.code_list.Add($line)
            if($controller.isend($line)){
                $controller = $controller.mother;
            }
        }
        return $this.children
    }
}









class Factory{
    $filepartslist = [System.Collections.ArrayList]::new()
    #$rootpath = "/Users/minegishirei/myworking/VBAToolKit/Source/ConfProd"
    $rootpath = "/Users/minegishirei/myworking/VBAToolKit/Source/ConfProd"

    [System.Collections.ArrayList]run(){
        $rootpathItem = Get-ChildItem -Recurse $this.rootpath
        foreach($item in $rootpathItem){
            if($item.PSIsContainer){
                #folder
            }else{
                #file
                $fileParts = [FileParts]::new()
                $fileParts.setName($item.Name)

                $filepath = [string]$item.Directory + "/"+[string]$item.Name
                $contextFactory  = [ContextFactory]::new()
                $contextFactory.setName($item.Name)
                $fileParts.children = $contextFactory.run($filepath)

                [void]$this.filepartslist.Add($fileParts)
            }
        }
        return $this.filepartslist
    }

    buildXML([string]$xmlFile){

        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
        
        [System.Xml.XmlDocument]$doc = [System.Xml.XmlDocument]::new()
        
        $dec = $doc.CreateXmlDeclaration("1.0", $null, $null)
        $doc.AppendChild($dec) | Out-Null
        
        $container = $doc.CreateNode("element", "container", $xmlns)
        $container.SetAttribute("version", "1.0")
        
        ##ここから中身
        $fileParts = $this.filepartslist[0]
        $fileParts.AppendXmlChild($doc, $doc)


        $doc.Save($xmlFile) | Out-Null

    }
}

$factory = New-Object Factory

$factory.run()

$factory.buildXML("/Users/minegishirei/myworking/ReadableChart/main/test.xml")
















