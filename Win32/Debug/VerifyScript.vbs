Option Explicit 'Делает текст VBScript более строгим. Объявление переменных перед их использованием становится обязательным

'====== Блок объявления переменных ======

Dim oPKCS7Message 'Создаём объект сообщения формата PKCS7 (специальный формат для работы с подписью)

Const DT_SIGNED_DATA = 2 'Криптографическое сообщение, содержащее результат формирования ЭЦП
Const CERT_AND_SIGN = 0 'Проверка подписи и сертификата
Const SIGN_ONLY = 1 'Проверка только подписи
Dim oSignatures 'Объект, содержащий коллекцию из объектов подписи (внутри файла *.sig может быть несколько подписей)
Dim oSignature 'Объект одной подписи
Dim Status
Dim oCertificate
Dim oCertificates : Set oCertificates = CreateObject("DigtCrypto.Certificates") 'Коллекция сертификатов. Переменная создана только для того, чтобы поместить в неё
                  'один сертификат. Без коллекции нельзя настроить профиль
Const POLICY_TYPE_NONE = 0 'Нет политики использования сертификатов
'____ Статусы сертификата ____
Const VS_CORRECT = 1
Const VS_UNSUFFICIENT_INFO = 2
Const VS_UNCORRECT = 3
Const VS_INVALID_CERTIFICATE_BLOB = 4
Const VS_CERTIFICATE_TIME_EXPIRIED = 5
Const VS_CERTIFICATE_NO_CHAIN  = 6
Const VS_CERTIFICATE_CRL_UPDATING_ERROR = 7
Const VS_LOCAL_CRL_NOT_FOUND = 8
Const VS_CRL_TIME_EXPIRIED = 9
Const VS_CERTIFICATE_IN_CRL = 10
Const VS_CERTIFICATE_IN_LOCAL_CRL = 11
Const VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL = 12
Const VS_CERTIFICATE_USING_RESTRICTED = 13

Dim oProfile : Set oProfile = CreateObject("DigtCrypto.Profile") 'Создаём профиль. Только внутри него можно выбрать виды проверок сертификата с CRL
Const CERTIFICATE_VERIFY_REV_PROV = 4 'Проверка сертификата с помощью Revocation Provider
Const CERTIFICATE_VERIFY_ONLINE_CRL = 2 'Проверка сертификата онлайн
Const CERTIFICATE_VERIFY_OCSP = 8 'Проверка сертификат в службе OCSP

Dim i

'====== Основной блок ======

Function SignatureVerify (ByVal InputFileName, ByVal InputFileNameSignature)
    Set oPKCS7Message = CreateObject("DigtCrypto.PKCS7Message")
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, InputFileName
    Set oSignatures = oPKCS7Message.Signatures
    Dim n : n = oSignatures.Count
    Dim arrayResults()
    Redim Preserve arrayResults(n-1)
    For i = 0 To n-1
        Set oSignature = oSignatures.Item(i)
        Status = oSignature.Verify (CERT_AND_SIGN)
        arrayResults(i) = Status
    Next
    SignatureVerify = arrayResults 'Возвращаем результат проверки

    'Очищаем переменные
    Set oSignature = Nothing
    Set oSignatures = Nothing
    Set oPKCS7Message = Nothing
End Function

Function SignatureInformation (ByVal InputFileNameSignature)
    Set oPKCS7Message = CreateObject("DigtCrypto.PKCS7Message")
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim SigInfo
    Dim n : n = oSignatures.Count
    Dim arrayResults()
    Redim Preserve arrayResults(n-1)
    For i=0 to n-1
        Set oSignature = oSignatures.Item(i)

        'Получаем указатель на сертификат подписи
        SigInfo =           "Подпись для сертификата № " + CStr(oSignature.Certificate.SerialNumber) + ":" + vbCrLf
        'Время подписи
        dim sSigningTime : sSigningTime = oSignature.SigningTime
        SigInfo = SigInfo + "Время подписи: " + CStr(sSigningTime) + vbCrLf
        'Хэш алгоритм подписи
        Dim sHashAlg : sHashAlg = oSignature.HashAlg
        SigInfo = SigInfo + "Хэш алгоритм подписи: " + CStr(sHashAlg) + vbCrLf
        'Алгоритм подписи ЭЦП
        Dim sHashEncAlg : sHashEncAlg = oSignature.HashEncAlg
        SigInfo = SigInfo + "Алгоритм подписи ЭЦП: " + CStr(sHashEncAlg) + vbCrLf
        'Тип ЭЦП: true – отсоединённая, false – присоединённая
        Dim sDetached : sDetached = oSignature.Detached
        Dim sDetachedValue        
        if sDetached then
            sDetachedValue = "Отсоединённая"
        else sDetachedValue = "Присоединённая"
        end if
        SigInfo = SigInfo + "Тип ЭЦП: " + sDetachedValue + vbCrLf
        'Номер версии протокола версии CMS
        Dim lCMSVersion : lCMSVersion = oSignature.CMSVersion
        SigInfo = SigInfo + "Номер версии протокола CMS: " + CStr(lCMSVersion) + vbCrLf + vbCrLf
        arrayResults(i) = SigInfo
    Next
    SignatureInformation = arrayResults 'Возвращаем результат информации о подписи

    'Очищаем переменные
    Set oSignature = Nothing
    Set oSignatures = Nothing
    Set oPKCS7Message = Nothing
End Function

Function CertificateInformation (ByVal InputFileNameSignature)
    Set oPKCS7Message = CreateObject("DigtCrypto.PKCS7Message")
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim CertInfo
    Dim n : n = oSignatures.Count
    Dim arrayResults()
    Redim Preserve arrayResults(n-1)
    For i=0 to n-1
        Set oSignature = oSignatures.Item(i)

        'Получаем указатель на сертификат подписи
        Set oCertificate = oSignature.Certificate

        CertInfo =            "Сертификат № " + CStr(oCertificate.SerialNumber) + ":" + vbCrLf + vbCrLf
        'Свойства сертификата
        CertInfo = CertInfo + "Выдан УЦ: " + CStr(oCertificate.IssuerName) + vbCrLf + vbCrLf
        CertInfo = CertInfo + "Владелец сертификата: " + CStr(oCertificate.SubjectName) + vbCrLf + vbCrLf
        CertInfo = CertInfo + "Действует с: " + CStr(oCertificate.ValidFrom) + vbCrLf
        CertInfo = CertInfo + "Действует по: " + CStr(oCertificate.ValidTo) + vbCrLf + vbCrLf
        arrayResults(i) = CertInfo
    Next
    CertificateInformation = arrayResults 'Возвращаем результат информации о сертификате

    'Очищаем переменные
    Set oCertificate = Nothing
    Set oSignature = Nothing
    Set oSignatures = Nothing
    Set oPKCS7Message = Nothing
End Function

Function CertificateVerify (ByVal InputFileNameSignature)
    Set oPKCS7Message = CreateObject("DigtCrypto.PKCS7Message")
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim CertVerify
    Dim n : n = oSignatures.Count
    Dim arrayResults()
    Redim Preserve arrayResults(n-1)
    For i=0 to n-1
        Set oSignature = oSignatures.Item(i)

        'Получаем указатель на сертификат подписи
        Set oCertificate = oSignature.Certificate
        oCertificates.Add oSignature.Certificate
        'Проверка статуса сертификата
        oProfile.SetVerifiedCertificates CERTIFICATE_VERIFY_REV_PROV, oCertificates
        'Установим профиль в сертификат и выполним проверку
        oCertificate.Profile = oProfile
        Status = -1
        Status = oCertificate.IsValid(POLICY_TYPE_NONE) 'Проверим статус сертификата

        Select Case Status
            Case VS_CORRECT
                CertVerify = "Статус сертификата: " + "Корректен"
            Case VS_UNSUFFICIENT_INFO
                CertVerify = "Статус сертификата: " + "Статус неизвестен"
            Case VS_UNCORRECT
                CertVerify = "Статус сертификата: " + "Некорректен"
            Case VS_INVALID_CERTIFICATE_BLOB
                CertVerify = "Статус сертификата: " + "Недействительный блоб сертификата"
            Case VS_CERTIFICATE_TIME_EXPIRIED
                CertVerify = "Статус сертификата: " + "Время действия сертификата истекло или ещё не наступило"
            Case VS_CERTIFICATE_NO_CHAIN
                CertVerify = "Статус сертификата: " + "Невозможно построить цепочку сертификации"
            Case VS_CERTIFICATE_CRL_UPDATING_ERROR
                CertVerify = "Статус сертификата: " + "Ошибка обновления сертификата"
            Case VS_LOCAL_CRL_NOT_FOUND
                CertVerify = "Статус сертификата: " + "Не найден локальный СОС"
            Case VS_CRL_TIME_EXPIRIED
                CertVerify = "Статус сертификата: " + "Истекло время действия СОС"
            Case VS_CERTIFICATE_IN_CRL
                CertVerify = "Статус сертификата: " + "Сертификат находится в СОС"
            Case VS_CERTIFICATE_IN_LOCAL_CRL
                CertVerify = "Статус сертификата: " + "Сертификат находится в локальном СОС"
            Case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL
                CertVerify = "Статус сертификата: " + "Сертификат действителен по локальному СОС"
            Case VS_CERTIFICATE_USING_RESTRICTED
                CertVerify = "Статус сертификата: " + "Использование сертификата ограничено"
        End Select

        arrayResults(i) = CertVerify
    Next
    CertificateVerify = arrayResults

    'Очищаем переменные
    Set oCertificate = Nothing
    Set oSignature = Nothing
    Set oSignatures = Nothing
    Set oPKCS7Message = Nothing
End Function