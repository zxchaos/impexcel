
@echo off
REM This script will add impexcel jars to your classpath.

set IMP_HOME=D:\impExcelExe\impExcel
echo %IMP_HOME%

set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\commons-lang3-3.1.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\dom4j-1.6.1.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\jaxen-1.1.4.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\poi-3.9-20121203.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\poi-ooxml-3.9-20121203.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\poi-ooxml-schemas-3.9-20121203.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\xbean.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\ojdbc14.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\commons-io-2.4.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\logback-access-1.0.13.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\logback-classic-1.0.13.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\logback-core-1.0.13.jar
set CLASSPATH=%CLASSPATH%;%IMP_HOME%\lib\slf4j-api-1.7.5.jar
echo %CLASSPATH%
cd %IMP_HOME%\bin
java  com.taiji.excelimp.util.RegionUtil