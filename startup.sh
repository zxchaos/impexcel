#!/bin/bash
echo "start up impExcel"
IMP_HOME=/home/zxchaos/projectFiles/projects/impExcel
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/commons-lang3-3.1.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/dom4j-1.6.1.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/jaxen-1.1.4.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/poi-3.9-20121203.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/poi-ooxml-3.9-20121203.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/poi-ooxml-schemas-3.9-20121203.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/xbean.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/ojdbc14.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/commons-io-2.4.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/logback-access-1.0.13.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/logback-classic-1.0.13.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/logback-core-1.0.13.jar
CLASSPATH=$CLASSPATH:$IMP_HOME/lib/slf4j-api-1.7.5.jar
CLASSPATH=$CLASSPATH%:$IMP_HOME/lib/commons-codec-1.9.jar
export IMP_HOME
export CLASSPATH
echo $CLASSPATH
cd $IMP_HOME/bin
java -Xms1000M -Xmx1600M -XX:PermSize=256M -XX:MaxPermSize=512M   com.taiji.excelimp.ImpCheck

