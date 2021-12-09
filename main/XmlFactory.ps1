
#Param($Path)

$Global:XMLDOC = [System.Xml.XmlDocument]::new()
$Global:XMLS = "urn:oasis:names:tc:opendocument:xmlns:container"



function is_safe_string([string]$inspect_str){
    $check_list = @(";", '"', " ", ",")
    foreach ($item in $check_list) {
        if($inspect_str.Contains($item)){
            return $False
        }
    }
    return $True
}


############################################
##Parts:####################################
############################################
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
    [System.Xml.XmlElement]create_element([string]$tag, [string]$innerText){
        $xmlElement = $Global:XMLDOC.CreateNode("element", $tag, $Global:XMLS )
        $xmlElement.InnerText = $innerText
        return $xmlElement
    }    
}

class FolderParts : Parts{
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

    [void]buildXML([string]$xmlFile){
        $xmlns = "urn:oasis:names:tc:opendocument:xmlns:container"
        $dec = $Global:XMLDOC.CreateXmlDeclaration("1.0", $null, $null)
        $Global:XMLDOC.AppendChild($dec) | Out-Null
        
        [System.Xml.XmlElement]$container = $Global:XMLDOC.CreateNode("element", "container", $xmlns)
        $container.SetAttribute("version", "1.0")
        ##ここから中身
        foreach($fileParts in $this.filepartslist){
            $fileParts.AppendXmlChild($container)
        }
        $Global:XMLDOC.AppendChild($container)
        ##ここまで中身
        $Global:XMLDOC.Save($xmlFile) | Out-Null
    }

    [string]buildUML([string]$inputXmlFile, [string]$outputUmlFile){
        $XmlObj = [System.Xml.XmlDocument](Get-Content $inputXmlFile) 
        [UMLFactory]$umlFactory = [UMLFactory]::new();
        $umlFactory.buildFileXMLloop($XmlObj.container);
        return "@startuml`n" + $umlFactory.umltext + "@enduml`n"
    }
}

class FileParts : Parts{
    $type = "file"
    AppendXmlChild($container){
        $child_xml = $this.create_element("child", "")
        $child_xml.AppendChild( $this.create_element("type", $this.type)  )
        $child_xml.AppendChild( $this.create_element("name", $this.name)  )
        $container.AppendChild($child_xml) | Out-Null
        foreach ($item in $this.children) {
            $item.AppendXmlChild($child_xml)
        }
    }

    [void] setName($filename){
        $this.name = $filename
    }
}



##TODO:global =>static変数
$Global:NAMESPACE = [System.Collections.ArrayList]::new();
class ContextParts : Parts {
    $code_list = [System.Collections.ArrayList]::new()

    [bool]isstart([string]$line){
        if( ($line -match $this.start_regex) -and (is_safe_string($Matches.match_name)) ){
            $this.name = $Matches.match_name
            $Global:NAMESPACE.Add($this.name)
            return $true
        }else{
            return $false
        }
    }

    [bool]isend([string]$line){
        return $line -match $this.end_regex
    }
    
    AppendXmlChild($parents_xml){
        $child_xml = $this.create_element("child", "")
        $child_xml.AppendChild( $this.create_element("type", $this.type)  )
        $child_xml.AppendChild( $this.create_element("name", $this.name)  )
        $child_xml.AppendChild( $this.create_element("code_list", $this.code_list -join "`n")  )
        ##check call namespace
        foreach ($line in $this.code_list) {
            foreach ($callablename in $Global:NAMESPACE) {
                if($line.contains(" "+$callablename+" ")){
                    $child_xml.AppendChild( $this.create_element("call", $callablename)  )
                }
            }
        }
        $parents_xml.AppendChild($child_xml) | Out-Null
        
        foreach ($item in $this.children) {
            if($item.type -eq "comment"){
                continue
            }
            $item.AppendXmlChild($child_xml)
        }
    }
}


class FunctionParts : ContextParts {
    [string]$type = "function";
    [string]$start_regex = "Function (?<match_name>.*?)\("
    [string]$end_regex = ".*End.*Function.*"
}


class SubParts : ContextParts {
    [string]$type = "sub";
    [string]$start_regex = "Sub (?<match_name>.*?)\("
    [string]$end_regex = ".*End.*Sub.*"
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


############################################
##ContextFactory:###########################
############################################
class ContextFactory : ContextParts{
    ##なぜか使われない
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
        $this.mother = $this
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
                $functionParts.mother = $controller
                $controller = $functionParts
            }

            $subParts = [SubParts]::new();
            if($subParts.isstart($line)){
                $controller.addChild($subParts)
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
    $src_text = ""
    buildFileXMLloop([System.Xml.XmlElement]$XmlObj){
        #file単位
        foreach ($fileitem in $XmlObj.ChildNodes) {
            $this.src_text = ""
            $filename = $fileitem.name
            #xmlタグ単位
            foreach ($item in $fileitem) {
                if( ($item.LocalName -eq "child")){
                    $this.buildXMLloop($item)
                }
            }
            $this.umltext += ( "file " + $filename)
            $this.umltext += ("{`n")
                $this.umltext += $this.src_text
            $this.umltext += ("}`n")
        }
    }

    buildXMLloop([System.Xml.XmlElement]$XmlObj){
        ##definition
        $this.src_text += ( "class " + $XmlObj.Name + " {`n")
            $this.src_text += ( "- " + $XmlObj.type +"`n")
        $this.src_text += ("}`n") 
        #$this.umltext += "note top`n"
        #    $src = $XmlObj.code_list -join "`n"
        #    $this.umltext += ($src)
        #$this.umltext += "end note`n"

        foreach ($item in $XmlObj.ChildNodes) {
            if( ($item.LocalName -eq "child")){
                $this.buildXMLloop($item)
            }elseif($item.LocalName -eq "call"){
                if($XmlObj.Name  -eq  $item.InnerText){
                    continue
                }
                $this.src_text += ( $XmlObj.Name + " -down-|> " + $item.InnerText ) +"`n"
            }
        }
    }
}






$factory = [FolderParts]::new()
[void]$factory.run( "/Users/minegishirei/myworking/VBAToolKit/Source/ConfProd")
$Global:NAMESPACE = $Global:NAMESPACE | Select-Object -Unique 
[void]$factory.buildXML("/Users/minegishirei/myworking/ReadableChart/src.xml")
Set-Clipboard  $factory.buildUML("/Users/minegishirei/myworking/ReadableChart/src.xml", "test.uml")

Read-Host "Exit"



