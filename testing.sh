remotepath=~/share/xml
file=test2.odt
xslt=odt-to-tagged2.xsl
result=results/test.txt
cp $remotepath/$file data
rm styles.xml
cd data
rm content.xml styles.xml
unzip $file content.xml styles.xml
cd ..
ln -s data/styles.xml
java -classpath ~/xml/saxon9he.jar net.sf.saxon.Transform data/content.xml $xslt -o:$result
# ~/py/recode.py -d UTF-16LE -r windows  $result $remotepath/result.txt
cp $result $remotepath/result.txt
