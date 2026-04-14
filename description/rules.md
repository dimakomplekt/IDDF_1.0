1) Никогда не использовать Set для объявления - ТОЛЬКО dim (Set — присвоение объектной переменной – не работает в инвенторе!!!). Объекты делае без 

``` bas

Set - fso = CreateObject("Scripting.FileSystemObject") ' Object, без Set

```

2) Не использовать команду либы Format() - менять на свои итераторы
   
3) НЕ ИСПОЛЬЗОВАТЬ КРУГЛЫЕ СКОБКИ НА КОНЦЕ ФУНКЦИЙ И САБОВ ПРИ ВЫЗОВЕ

``` bas

' Так нельзя!

oContext = app.TransientObjects.CreateTranslationContext()
oOptions = app.TransientObjects.CreateNameValueMap()
oData = app.TransientObjects.CreateDataMedium()

' Надо так!

oContext = app.TransientObjects.CreateTranslationContext
oOptions = app.TransientObjects.CreateNameValueMap
oData = app.TransientObjects.CreateDataMedium

```

4) Использовать круглые скобки при любой передаче параметров / аргументов

``` bas

' Так нельзя!
MsgBox log, vbInformation, "DXF Export Completed"

' Надо так!
MsgBox(log, vbInformation, "DXF Export Completed")

```

5) 'Тип "Variant" больше не поддерживается; используйте тип "Object". !!!