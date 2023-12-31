*** Settings ***
Documentation       Robot prueba formulario google
Library    RPA.Browser.Selenium    auto_close=${FALSE}
Library    RPA.Excel.Files
Library    RPA.PDF
Library    RPA.HTTP
Library    RPA.Desktop
Library    RPA.Excel.Application
Library    RPA.Word.Application
Library    BuiltIn
Library    RPA.Email.ImapSmtp    smtp_server=smtp.gmail.com    smtp_port=587
Task Setup    Authorize    account=%{EMAIL}    password=%{PASSWORD}

*** Variables ***
${RUTA_EXCEL}=    ${CURDIR}${/}Datos.xlsx${/}
${CARPETA}=    ${CURDIR}${/}archivos${/}
# ${CORREO}=    juanfran1019@hotmail.com

*** Tasks ***
Robot prueba formulario google
    Acceder al formulario
    Descargar Excel y leer columnas
    Cerrar navegador y formulario

*** Keywords ***
Acceder al formulario
    Open Chrome Browser    https://docs.google.com/forms/d/e/1FAIpQLScNyPg715I1DXzlb_W4i18YvHXHfGG3bpsMyfG-0wHQRCCc3g/viewform?usp=sf_link
    Maximize Browser Window
    Wait Until Element Is Visible    css:#mG61Hd > div.RH5hzf.RLS9Fe > div > div.Dq4amc > div > div.N0gd6 > div.ahS2Le > div
    Log    Entrada al formulario correcta

Introducir cada fila al formulario
    [Arguments]    ${fila}
    Input Text    xpath://html/body/div/div[2]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input    ${fila}[Nombre]
    Input Text    xpath://html/body/div/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input    ${fila}[Apellidos]
    Input Text    xpath://html/body/div/div[2]/form/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea    ${fila}[Dirección]
    Input Text    xpath://html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input    ${fila}[Edad]
    
    IF    "${fila}[Sexo]" == "M"
        Log    es un hombre
        Click Element    css:#i21 > div.vd3tt > div
    END
    IF    "${fila}[Sexo]" == "F"
        Log    es una mujer
        Click Element   css:#i24 > div.vd3tt > div
    END
    IF    "${fila}[Sexo]" == $null
        Log    prefiere no decir su sexo
        Click Element   css:#i27 > div.vd3tt > div
    END
    Click Element    css:#mG61Hd > div.RH5hzf.RLS9Fe > div > div.ThHDze > div.DE3NNc.CekdCb > div.lRwqcd > div > span
    
    Enviar un correo electrónico de confirmación    ${fila}

    Volver a rellenar otro formulario

Enviar un correo electrónico de confirmación
    [Arguments]    ${fila}
    ${contenido}    Set Variable    Nombre:${fila}[Nombre],${\n}${\n}Apellidos:${fila}[Apellidos],${\n}${\n}Edad:${fila}[Edad],${\n}${\n}Sexo:${fila}[Sexo],${\n}${\n}Dirección:${fila}[Dirección]
    Html To Pdf    content=${contenido}    output_path=${CARPETA}trabador_${fila}[Nombre].pdf
    Send Message    sender=%{EMAIL}    
    ...    recipients=cvivancos@cenitcon.com
    ...    subject=FORMULARIO RELLENADO
    ...    body=Buenos Días, ${fila}[Nombre] ${fila}[Apellidos]. Su formulario ha sido enviado con los datos mostrados en el PDF adjunto. Compruebe que son correctos y póngase en contacto con nosotros si no fuera así. Un saludo.
    ...    attachments=${CARPETA}trabador_${fila}[Nombre].pdf

Volver a rellenar otro formulario
    Wait Until Element Is Visible    css:body > div.Uc2NEf > div:nth-child(2) > div.RH5hzf.RLS9Fe > div > div.vHW8K
    Click Element    css:body > div.Uc2NEf > div:nth-child(2) > div.RH5hzf.RLS9Fe > div > div.c2gzEf > a
    Log    Se va a rellenar otro formulario

Descargar Excel y leer columnas
    RPA.Excel.Files.Open Workbook    ${RUTA_EXCEL}
    ${tabla}=    Read Worksheet As Table    header=${True}
    Close Workbook

    FOR    ${row}    IN    @{tabla}
        Log    ${row}
        Introducir cada fila al formulario    ${row}
    END

Cerrar navegador y formulario
    Close Window
    Log    Se ha cerrado la ventana