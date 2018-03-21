# java2excel

Tool for java reading/writing excel file.

## Install 
- clone project into local disk
```
git clone https://github.com/johnsonmoon/java2excel.git
```
- install apache maven, execute command below in the project directory 
```
mvn install -Dmaven.test.skip=true
```

## Include java2excel into your project path (maven)
- pom.xml
```
<dependency>
    <groupId>xuyihao</groupId>
    <artifactId>java2excel</artifactId>
    <version>${version}</version>
</dependency>

Choose your favourite ${version}.
```

## Using
There are a few ways to edit/write/read excel file with java2excel tool.

### Interface & Class

number | name | interface or class | description
--- | --- | --- | ---
1 | [Editor](readme/Editor.md) | interface | interface for edit excel file (including read & write)
2 | [Reader](readme/Writer.md) | interface | interface for read data from excel file
3 | [Writer](readme/Reader.md) | interface | interface for write data into excel file
 | | |
4 | [CustomEditor](readme/type-custom.md) | class | class implements Editor, custom editing excel file (including read & write)
5 | [CustomReader](readme/type-custom.md) | class | class implements Reader, custom reading data from excel file
6 | [CustomWriter](readme/type-custom.md) | class | class implements Writer, custom writing data into excel file
 | | |
7 | [FormattedEditor](readme/type-formatted.md) | class | class implements Editor, editing excel file with formatted way
8 | [FormattedReader](readme/type-formatted.md) | class | class implements Reader, reading data from excel file with formatted way
9 | [FormattedWriter](readme/type-formatted.md) | class | class implements Writer, writing data into excel file with formatted way
 | | |
10 | [FreeEditor](readme/type-free.md) | class | class editing excel file freely
11 | [FreeReader](readme/type-free.md) | class | class reading excel file freely
12 | [FreeWriter](readme/type-free.md) | class | class Writing excel file freely