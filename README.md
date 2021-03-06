officefilelocker
================

A Jython CLI Tool enables you to protect Office File Documents from tampering.

### Dependencies
 - [Java RE](https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html) - Java Runtime Engine
 - [Jython CLI](http://www.jython.org/) - Python for the Java Platform
 - [Apache POI](http://poi.apache.org/) -  Java API for Microsoft Documents

### Configuration
Copy the template config and add the Apache POI library to path

```bash
cp apache.cfg.example apache.cfg
# [POI]
# path: ~/packages/apache/poi
```
### Jython Evironment with pyenv
Consider using pyenv to isolate your jython environment

```bash
pyenv install jython-2.7.0 --keep --verbose
pyens shell jython-2.7.0
jython --version
```

### Usage
Locking an Office File:

```bash
jython officefilelocker.py -u <username> -p <password> -i <inputfile> -o <outputfile>
```

### Todo

- [ ] [Protect the XWPFDocument(.docx) with Password](http://poi.apache.org/apidocs/org/apache/poi/xwpf/usermodel/XWPFDocument.html#enforceReadonlyProtection(java.lang.String, org.apache.poi.poifs.crypt.HashAlgorithm))
- [ ] [Convert HWPFDocument(.doc) to XWPFDocument(.doxc) File Stream](http://poi.apache.org/apidocs/org/apache/poi/hwpf/HWPFDocument.html)

### License

[MIT](https://github.com/nathanielvarona/officefilelocker/blob/master/LICENSE)
