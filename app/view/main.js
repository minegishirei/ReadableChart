import {ReadableChart} from '../control/readablechart';



class FormMain{

    //propaty
    readablechart = ReadableChart()

    getfolderpath(){
        var element = document.getElementById("folderpath")
        var foloderpath = element.attributes["value"]
        return foloderpath
    }


    submit() {
        var folderpath = this.getfolderpath()
        readablechart.read()
        
    }
}


formMain = FormMain()


