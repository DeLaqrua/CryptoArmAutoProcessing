Option Explicit '������ ����� VBScript ����� �������. ���������� ���������� ����� �� �������������� ���������� ������������

'====== ���� ���������� ���������� ======

Dim oPKCS7Message '������ ������ ��������� ������� PKCS7 (����������� ������ ��� ������ � ��������)

Const DT_SIGNED_DATA = 2 '����������������� ���������, ���������� ��������� ������������ ���
Const CERT_AND_SIGN = 0 '�������� ������� � �����������
Const SIGN_ONLY = 1 '�������� ������ �������
Dim oSignatures '������, ���������� ��������� �� �������� ������� (������ ����� *.sig ����� ���� ��������� ��������)
Dim oSignature '������ ����� �������
Dim Status
Dim oCertificate
Dim oCertificates : Set oCertificates = CreateObject("DigtCrypto.Certificates") '��������� ������������. ���������� ������� ������ ��� ����, ����� ��������� � ��
                  '���� ����������. ��� ��������� ������ ��������� �������
Const POLICY_TYPE_NONE = 0 '��� �������� ������������� ������������
'____ ������� ����������� ____
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

Dim oProfile : Set oProfile = CreateObject("DigtCrypto.Profile") '������ �������. ������ ������ ���� ����� ������� ���� �������� ����������� � CRL
Const CERTIFICATE_VERIFY_REV_PROV = 4 '�������� ����������� � ������� Revocation Provider
Const CERTIFICATE_VERIFY_ONLINE_CRL = 2 '�������� ����������� ������
Const CERTIFICATE_VERIFY_OCSP = 8 '�������� ���������� � ������ OCSP

Dim i

'====== �������� ���� ======

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
    SignatureVerify = arrayResults '���������� ��������� ��������

    '������� ����������
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

        '�������� ��������� �� ���������� �������
        SigInfo =           "������� ��� ����������� � " + CStr(oSignature.Certificate.SerialNumber) + ":" + vbCrLf
        '����� �������
        dim sSigningTime : sSigningTime = oSignature.SigningTime
        SigInfo = SigInfo + "����� �������: " + CStr(sSigningTime) + vbCrLf
        '��� �������� �������
        Dim sHashAlg : sHashAlg = oSignature.HashAlg
        SigInfo = SigInfo + "��� �������� �������: " + CStr(sHashAlg) + vbCrLf
        '�������� ������� ���
        Dim sHashEncAlg : sHashEncAlg = oSignature.HashEncAlg
        SigInfo = SigInfo + "�������� ������� ���: " + CStr(sHashEncAlg) + vbCrLf
        '��� ���: true � ������������, false � �������������
        Dim sDetached : sDetached = oSignature.Detached
        Dim sDetachedValue        
        if sDetached then
            sDetachedValue = "������������"
        else sDetachedValue = "�������������"
        end if
        SigInfo = SigInfo + "��� ���: " + sDetachedValue + vbCrLf
        '����� ������ ��������� ������ CMS
        Dim lCMSVersion : lCMSVersion = oSignature.CMSVersion
        SigInfo = SigInfo + "����� ������ ��������� CMS: " + CStr(lCMSVersion) + vbCrLf + vbCrLf
        arrayResults(i) = SigInfo
    Next
    SignatureInformation = arrayResults '���������� ��������� ���������� � �������

    '������� ����������
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

        '�������� ��������� �� ���������� �������
        Set oCertificate = oSignature.Certificate

        CertInfo =            "���������� � " + CStr(oCertificate.SerialNumber) + ":" + vbCrLf + vbCrLf
        '�������� �����������
        CertInfo = CertInfo + "����� ��: " + CStr(oCertificate.IssuerName) + vbCrLf + vbCrLf
        CertInfo = CertInfo + "�������� �����������: " + CStr(oCertificate.SubjectName) + vbCrLf + vbCrLf
        CertInfo = CertInfo + "��������� �: " + CStr(oCertificate.ValidFrom) + vbCrLf
        CertInfo = CertInfo + "��������� ��: " + CStr(oCertificate.ValidTo) + vbCrLf + vbCrLf
        arrayResults(i) = CertInfo
    Next
    CertificateInformation = arrayResults '���������� ��������� ���������� � �����������

    '������� ����������
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

        '�������� ��������� �� ���������� �������
        Set oCertificate = oSignature.Certificate
        oCertificates.Add oSignature.Certificate
        '�������� ������� �����������
        oProfile.SetVerifiedCertificates CERTIFICATE_VERIFY_REV_PROV, oCertificates
        '��������� ������� � ���������� � �������� ��������
        oCertificate.Profile = oProfile
        Status = -1
        Status = oCertificate.IsValid(POLICY_TYPE_NONE) '�������� ������ �����������

        Select Case Status
            Case VS_CORRECT
                CertVerify = "������ �����������: " + "���������"
            Case VS_UNSUFFICIENT_INFO
                CertVerify = "������ �����������: " + "������ ����������"
            Case VS_UNCORRECT
                CertVerify = "������ �����������: " + "�����������"
            Case VS_INVALID_CERTIFICATE_BLOB
                CertVerify = "������ �����������: " + "���������������� ���� �����������"
            Case VS_CERTIFICATE_TIME_EXPIRIED
                CertVerify = "������ �����������: " + "����� �������� ����������� ������� ��� ��� �� ���������"
            Case VS_CERTIFICATE_NO_CHAIN
                CertVerify = "������ �����������: " + "���������� ��������� ������� ������������"
            Case VS_CERTIFICATE_CRL_UPDATING_ERROR
                CertVerify = "������ �����������: " + "������ ���������� �����������"
            Case VS_LOCAL_CRL_NOT_FOUND
                CertVerify = "������ �����������: " + "�� ������ ��������� ���"
            Case VS_CRL_TIME_EXPIRIED
                CertVerify = "������ �����������: " + "������� ����� �������� ���"
            Case VS_CERTIFICATE_IN_CRL
                CertVerify = "������ �����������: " + "���������� ��������� � ���"
            Case VS_CERTIFICATE_IN_LOCAL_CRL
                CertVerify = "������ �����������: " + "���������� ��������� � ��������� ���"
            Case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL
                CertVerify = "������ �����������: " + "���������� ������������ �� ���������� ���"
            Case VS_CERTIFICATE_USING_RESTRICTED
                CertVerify = "������ �����������: " + "������������� ����������� ����������"
        End Select

        arrayResults(i) = CertVerify
    Next
    CertificateVerify = arrayResults

    '������� ����������
    Set oCertificate = Nothing
    Set oSignature = Nothing
    Set oSignatures = Nothing
    Set oPKCS7Message = Nothing
End Function