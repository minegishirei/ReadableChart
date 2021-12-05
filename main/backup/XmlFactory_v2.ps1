$Global:NAMESPACE = [System.Collections.ArrayList]::new();

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

class FolderParts : Parts{}


class FileParts : Parts{
    $type = "file"
    [System.Xml.XmlDocument]AppendXmlChild($container, [System.Xml.XmlDocument]$doc){
        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
        $child_xml = $doc.CreateNode("element", "child", $xmlns)
            $name_xml = $doc.CreateNode("element", "name", $xmlns)
                $name_xml.InnerText = $this.name
            $child_xml.AppendChild($name_xml)
            $type_xml = $doc.CreateNode("element", "type", $xmlns)
                $type_xml.InnerText = $this.type
            $child_xml.AppendChild($type_xml)
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

    [System.Xml.XmlDocument]AppendXmlChild($container, [System.Xml.XmlDocument]$doc){
        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
        $child_xml = $doc.CreateNode("element", "child", $xmlns)
            $name_xml = $doc.CreateNode("element", "name", $xmlns)
                $name_xml.InnerText = $this.name
            $child_xml.AppendChild($name_xml)
            $type_xml = $doc.CreateNode("element", "type", $xmlns)
                $type_xml.InnerText = $this.type
            $child_xml.AppendChild($type_xml)
            $code_xml = $doc.CreateNode("element", "code_list", $xmlns)
                $code = "`n"
                foreach ($item in $this.code_list) {
                    $code += ($item+"`n")
                }
                $code_xml.InnerText = $code
            $child_xml.AppendChild($code_xml) | Out-Null
            ##check call namespace
            foreach ($line in $this.code_list) {
                foreach ($callablename in $Global:NAMESPACE) {
                    if($line.contains(" "+$callablename+" ")){
                        $call_xml = $doc.CreateNode("element", "call", $xmlns)
                        $call_xml.InnerText = $callablename
                        $child_xml.AppendChild($call_xml)
                    }
                }
            }
        $container.AppendChild($child_xml) | Out-Null
        
        ##子アイテム
        foreach ($item in $this.children) {
            if($item.type -eq "comment"){
                continue
            }
            $item.AppendXmlChild($child_xml, $doc)
        }
        return $doc
    }
}


class FunctionParts : ContextParts {
    [string]$type = "function";

    [bool]isstart([string]$line){
        return ( ($line -like "*Function *") -and ($line -notlike "*End*Function*" ) )
    }

    [bool]isend([string]$line){
        return $line -like "*End*Function*"
    }
    
    [void]setName([string]$functionline){
        try{
            $right = $functionline.Split("Function")[1]
            $middle = $right.Split("(")[0]
            $this.name = $middle.Trim()
            $Global:NAMESPACE.Add($this.name)
        }catch{
            Write-Output "Something threw an exception or used Write-Error"
            Write-Output $_
        }
    }
}

class SubParts : ContextParts {
    [string]$type = "sub";

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
            $Global:NAMESPACE.Add($this.name)
        }catch{
            Write-Output "Something threw an exception or used Write-Error"
            Write-Output $_
        }
    }
}

class CommentParts : ContextParts {
    [string]$type = "comment";

    [bool]isstart([string]$line){
        return ( $line.Trim().StartsWith("'") )
    }

    [bool]isend([string]$line){
        return -not ( $line.Trim().StartsWith("'") )
    }
    
    [void]setName([string]$commentline){
        $this.name = ""
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
            $commentParts = [CommentParts]::new();
            if($commentParts.isstart($line)){
                continue
            }

            $functionParts = [FunctionParts]::new();
            if($functionParts.isstart($line)){
                $controller.addChild($functionParts)
                $functionParts.setName($line)
                $functionParts.mother = $controller
                $controller = $functionParts
            }

            $subParts = [SubParts]::new();
            if($subParts.isstart($line)){
                $controller.addChild($subParts)
                $subParts.setName($line)
                $subParts.mother = $controller
                $controller = $subParts
            }

            
            $controller.code_list.Add($line)
            if($controller.isend($line)){
                $controller = $controller.mother;
            }
        }
        return $this.children
    }
}




class UMLFactory{
    $umltext = ""
    buildXMLloop([System.Xml.XmlElement]$XmlObj){
        ##definition
        $this.umltext += ( "class " + $XmlObj.Name + " {}") +"`n"
        ##call
        #foreach ($item in $XmlObj) {
        #    if($item.Name -eq "call"){
        #        $this.umltext += ( $XmlObj.Name + " --|> " + $item.Name ) +"`n"
        #    }
        #}
        ##child
        foreach ($item in $XmlObj.ChildNodes) {
            if ($item.LocalName -eq "type") {
            }
            if( ($item.LocalName -eq "child")){
                $this.buildXMLloop($item)
            }elseif($item.LocalName -eq "call"){
                $this.umltext += ( $XmlObj.Name + " --|> " + $item.InnerText ) +"`n"
            }
        }
    }
}


class Factory{
    $filepartslist = [System.Collections.ArrayList]::new()
    [System.Collections.ArrayList]run([string]$rootpath){
        $rootpathItem = Get-ChildItem -Recurse $rootpath
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
        ##ここまで中身
        $doc.Save($xmlFile) | Out-Null
    }

    [string]buildPlantUML([string]$inputXmlFile, [string]$outputUmlFile){
        $XmlObj = [System.Xml.XmlDocument](Get-Content $inputXmlFile) 
        [UMLFactory]$umlFactory = [UMLFactory]::new();
        $umlFactory.buildXMLloop($XmlObj.child);
        return $umlFactory.umltext
    }

}






$factory = New-Object Factory
[void]$factory.run( "/Users/minegishirei/myworking/VBAToolKit/Source/ConfProd/")
[void]$factory.buildXML("/Users/minegishirei/myworking/ReadableChart/main/src.xml")
$factory.buildPlantUML("/Users/minegishirei/myworking/ReadableChart/main/src.xml", "test.uml")


