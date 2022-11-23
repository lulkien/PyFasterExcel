import os
from typing import Type

# TsAttribute class
class TsAttribute:
    def __init__(self, name: str, value: str) -> None:
        if not type(name) is str or not type(value) is str:
            raise TypeError("name or value must be string")
        self.__name = name
        self.__value = value

    def name(self) -> str:
        return self.__name

    def value(self) -> str:
        return self.__value

# TsWriter class
class TsWriter:
    def __init__(self, toFile: str) -> None:
        if not toFile:
            os.abort()

        os.makedirs(os.path.dirname(toFile), exist_ok = True)
        # open a file to write
        self.__file = open(toFile, "w")
        self.__initData()
    
    def __del__(self):
        while self.__tagStack:
            self.closeTag()
        self.__file.close()
    
    # Private methods
    def __initData(self) -> None:
        self.__tabWidth = 2
        self.__depth = 0
        self.__tagStack = []
        self.__attributes = []

    def __indent(self) -> None:
        self.__file.write(' ' * (self.__tabWidth * self.__depth))

    def __increaseDepth(self) -> None:
        self.__depth = self.__depth + 1
    
    def __decreaseDepth(self) -> None:
        if self.__depth == 0:
            return
        self.__depth = self.__depth - 1
    
    def __push(self, data: str) -> None:
        self.__tagStack.append(data)

    def __pop(self) -> str:
        retVal = self.__tagStack.pop()
        return retVal

    def __writeFile(self, data: str) -> None:
        self.__file.write(data)

    def __appendAttribute(self, attri: TsAttribute) -> None:
        self.__attributes.append(attri)
    
    def __clearAttribute(self) -> None:
        self.__attributes.clear()

    # Public methods
    def writeHeader(self) -> None:
        self.__file.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
        self.__file.write("<!DOCTYPE TS[]>\n")

    def openTag(self, tagName: str) -> None:
        if not type(tagName) is str:
            raise TypeError("Tag Name must be a string")
        self.__indent()
        self.__writeFile("<{tag}>\n".format(tag=tagName))
        self.__increaseDepth()
        self.__push(tagName)

    def openInlineTag(self, tagName: str, data: str) -> None:
        if not (type(tagName) is str and type(data) is str):
            raise TypeError("Tag Name and Data must be strings")
        self.__indent()
        strData = "<{tag}>{content}</{tag}>\n".format(tag=tagName, content=data)
        self.__writeFile(strData)

    def closeTag(self) -> None:
        if not self.__tagStack:
            print("Tags stack is empty")
            return
        self.__decreaseDepth()
        self.__indent()
        tagName = self.__pop()
        self.__writeFile("</{tag}>\n".format(tag=tagName))

    def addAttribute(self, name: str, value: str) -> None:
        if not (type(name) is str and type(value) is str):
            raise TypeError("Attribute name and value must be string")
        attri = TsAttribute(name, value)
        self.__appendAttribute(attri)

    def addAttributes(self, names, values):
        min_size = len(names) if len(names) < len(values) else len(values)
        for i in range(0, min_size):
            attri = TsAttribute(names[i], values[i])
            self.__appendAttribute(attri)

    def openTagWithAttribute(self, tagName: str) -> None:
        if not self.__attributes:
            self.openTag(tagName)
        if not type(tagName) is str:
            raise TypeError("Tag name must be string")
        self.__indent()
        strData = "<" + tagName
        for item in self.__attributes:
            strData = strData + " " + item.name() + "=\"" + item.value() + "\""
        strData += ">\n"
        self.__writeFile(strData)
        self.__increaseDepth()
        self.__push(tagName)
        self.__clearAttribute()

