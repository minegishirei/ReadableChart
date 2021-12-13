
#Param($Path)
##TODO:global =>static変数
$Global:NAMESPACE = [System.Collections.ArrayList]::new();
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



class UMLLib{
    [string]startUML(){
        return "@startuml`n"
    }

    [string]endUML(){
        return "@enduml`n"
    }
    
    [string]getConfig($config_text){
        return ($config_text + "`n")
    }

    [string]getFile($filename, $inside){
        $src_text = ( "file " + $filename)
        $src_text += ("{`n")
            $src_text += $inside
        $src_text += ("}`n")
        return $src_text
    }

    [string]getClass($name, $type){
        $src_text += ( "class " + $name + " {`n")
            $src_text += ( "- " + $type +"`n")
        $src_text += ("}`n") 
        return $src_text
    }

    [string]getClassAndDraw($name, $type, $color){
        $src_text += ( "class " + $name + " " + $color +" {`n")
            $src_text += ( "- " + $type +"`n")
        $src_text += ("}`n") 
        return $src_text
    }

    [string]getArrow($from, $dest){
        return ($from + " -down-|> " + $dest + "`n")
    }

    [string]getNote($name, $note){
        $src_text += "`n note top of $name `n"
            $src_text += ($note) + "`n"
        $src_text += "end note`n"
        return $src_text
    }

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
    [System.Xml.XmlElement]create_element([string]$tag, [string]$innerText){
        $xmlElement = $Global:XMLDOC.CreateNode("element", $tag, $Global:XMLS )
        $xmlElement.InnerText = $innerText
        return $xmlElement
    }    
}


##TODO:綺麗に
class FolderParts : Parts{
    $children = [System.Collections.ArrayList]::new()

    [void]run([string]$rootpath){
        $rootpathItem = Get-ChildItem -Recurse $rootpath
        foreach($fileitem in $rootpathItem){
            if($fileitem.PSIsContainer){
                #folder
            }else{
                #file
                $filepath = [string]$fileitem.Directory + "/"+[string]$fileitem.Name
                $fileParts = [FileParts]::new()
                $fileParts.setName($fileitem.Name)
                $fileParts.run($filepath)
                [void]$this.children.Add($fileParts)
            }
        }
    }

    [void]buildXML([string]$xmlFile){
        [System.Xml.XmlElement]$container = $Global:XMLDOC.CreateNode("element", "container", $Global:XMLS)
        $container.SetAttribute("version", "1.0")
        ##ここから中身
        foreach($fileParts in $this.children){
            $fileParts.AppendXmlChild($container)
        }
        $Global:XMLDOC.AppendChild($container)
        ##ここまで中身
        $Global:XMLDOC.Save($xmlFile) | Out-Null
    }
}



class FileParts : Parts{
    $type = "file"
    $code_list = [System.Collections.ArrayList]::new()

    [void] setName($filename){
        $this.name = $filename
    }

    [bool]isend($line){
        return $false
    }
    [bool]isstart($line){
        return $True
    }

    AppendXmlChild($container){
        $child_xml = $this.create_element("child", "")
        $child_xml.AppendChild( $this.create_element("type", $this.type)  )
        $child_xml.AppendChild( $this.create_element("name", $this.name)  )
        $container.AppendChild($child_xml) | Out-Null
        foreach ($item in $this.children) {
            $item.AppendXmlChild($child_xml)
        }
    }

    run($filepath){
        ## TODO:なくす
        $this.mother = $this; 
        $controller = $this;
        $contents = (Get-Content $filepath)

        $i = 0;
        foreach ($line in $contents) {
            $i += 1;
            ## TODO:loop instance and find appropriate one
            $commentParts = [CommentParts]::new();
            if($commentParts.isstart($line)){
                $controller.addChild($commentParts)
                $commentParts.mother = $controller
                $controller = $commentParts
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
    }
}



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
        #$child_xml.AppendChild( $this.create_element("comment", $this.)  )
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
    [string]$start_regex = "Function (?<match_name>.+?)\("
    [string]$end_regex = ".*End.*Function.*"
}



class SubParts : ContextParts {
    [string]$type = "sub";
    [string]$start_regex = "Sub (?<match_name>.+?)\("
    [string]$end_regex = ".*End.*Sub.*"
}



class CommentParts : ContextParts {
    [string]$type = "comment";
    [string]$start_regex = "^(\s)+'"

    [bool]isstart([string]$line){
        return ($line -match $this.start_regex)
    }

    [bool]isend($line){
        return $True
        #return -not ($line -match $this.start_regex)
    }
}



##TODO:なくす⇨Partsに移植。
class UMLFactory{
    $uml = [UMLLib]::new()
    $config_text = $this.uml.startUML()
    
    SetUMLConfig([string]$config_text){
        $this.config_text += $this.uml.getConfig($config_text)
    }

    [string]buildUML([string]$inputXmlFile){
        $XmlObj = [System.Xml.XmlDocument](Get-Content $inputXmlFile);
        $umltext = $this.buildFileXMLloop($XmlObj.container);
        $umltext = ($this.config_text + $umltext)
        return $umltext;
    }

    [string]buildFileXMLloop([System.Xml.XmlElement]$XmlObj){
        #file単位
        $umltext = ""
        foreach ($fileitem in $XmlObj.ChildNodes) {
            $filename = $fileitem.name
            #xmlタグ単位
            $context_src = ""
            foreach ($item in $fileitem.ChildNodes) {
                if( $item.LocalName -eq "child" ){
                    $context_src += $this.buildXMLloop($item)
                }
            }
            $umltext += $this.uml.getFile($filename, $context_src)
        }
        return  $umltext + $this.uml.endUML()
    }

    [string]buildXMLloop([System.Xml.XmlElement]$ContextPartsXml){
        ##definition
        $src_text = ""
        foreach ($item in $ContextPartsXml.ChildNodes) {

            if( ($item.LocalName -eq "child")){ #関数内関数定義
                $this.buildXMLloop($item)
            }
            elseif($item.LocalName -eq "code_list"){
                if($item.InnerText.Contains("sql")){
                    $src_text += $this.uml.getClassAndDraw($ContextPartsXml.Name, $ContextPartsXml.type, "#1e90ff")
                }else{
                    $src_text += $this.uml.getClass($ContextPartsXml.Name, $ContextPartsXml.type)
                }
                #$src_text += $this.uml.getNote($ContextPartsXml.Name, $item.InnerText)
            }
            elseif($item.LocalName -eq "call"){
                if($ContextPartsXml.Name  -eq  $item.InnerText){
                    continue
                }
                $src_text += $this.uml.getArrow( $ContextPartsXml.Name ,$item.InnerText )
            }
        }
        return $src_text
    }
}






## variable.
$target_file_path   = "./test/testsrc1/vtkReferenceManager.cls"
$xml_file_path      = "./material/material.xml"

## adbance ver.
#$target_file_path   = "./test/testsrc2/"
#$xml_file_path      = "./material/material.xml"

## building xml.
$factory = [FolderParts]::new()
[void]$factory.run($target_file_path)
$Global:NAMESPACE = $Global:NAMESPACE | Select-Object -Unique 
[void]$factory.buildXML($xml_file_path)

## reading  xml and building uml.
[UMLFactory]$umlFactory = [UMLFactory]::new();
$umlFactory.SetUMLConfig("!define DARKBLUE")
$umlFactory.SetUMLConfig("!include https://raw.githubusercontent.com/Drakemor/RedDress-PlantUML/master/style.puml")

$umlFactory.buildUML($xml_file_path) > src.uml
#Set-Clipboard | $umlFactory.buildUML($xml_file_path)


