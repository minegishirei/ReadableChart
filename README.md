# ReadableChart(what is this?)

you can automatically build "call graph" from source code written in vba and vb6.

<img src="https://raw.githubusercontent.com/kawadasatoshi/ReadableChart/main/sample/sample.svg">

This original srouce code is <a href="https://github.com/jpimbert/VBAToolKit/blob/master/Source/ConfProd/vtkReferenceManager.cls">VBAToolKit</a>.



## how to install

run next command in powershell terminal.

<pre><code>
git clone git@github.com:kawadasatoshi/ReadableChart.git.

</code></pre>

## how to run

run next command and open src.svg.

<pre><code>
./main.ps1              #build src.xml and src.uml
plantuml -svg src.uml   #build src.svg

</code></pre>






