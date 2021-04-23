Option Explicit '������ ����� VBScript ����� �������. ���������� ���������� ����� �� �������������� ���������� ������������

'====== ���� ���������� ���������� ======

Dim oPKCS7Message : Set oPKCS7Message = CreateObject("DigtCrypto.PKCS7Message") '������ ������ ��������� ������� PKCS7 (����������� ������ ��� ������ � ��������)

Const DT_SIGNED_DATA = 2 '����������������� ���������, ���������� ��������� ������������ ���
Const CERT_AND_SIGN = 0 '�������� ������� � �����������
Const SIGN_ONLY = 1 '�������� ������ �������
Dim oSignatures '������, ���������� ��������� �� �������� �������
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
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, InputFileName
    Set oSignatures = oPKCS7Message.Signatures
    Dim n : n = oSignatures.Count
    For  i=0 To n-1
        Set oSignature = oSignatures.Item(i)
        Status = oSignature.Verify (CERT_AND_SIGN)
        SignatureVerify = Status '���������� ��������� ��������       
    Next

    '������� ����������
    Set oSignature = Nothing
    Set oSignatures = Nothing
End Function

Function SignatureInformation (ByVal InputFileNameSignature)
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim n : n = oSignatures.Count
    For i=0 to n-1
        Set oSignature = oSignatures.Item(i)

        '����� �������
        dim sSigningTime : sSigningTime = oSignature.SigningTime
        SignatureInformation = SignatureInformation + "����� �������: " + CStr(sSigningTime) + vbCrLf
        '��� �������� �������
        Dim sHashAlg : sHashAlg = oSignature.HashAlg
        SignatureInformation = SignatureInformation + "��� �������� �������: " + CStr(sHashAlg) + vbCrLf
        '�������� ������� ���
        Dim sHashEncAlg : sHashEncAlg = oSignature.HashEncAlg
        SignatureInformation = SignatureInformation + "�������� ������� ���: " + CStr(sHashEncAlg) + vbCrLf
        '��� ���: true � ������������, false � �������������
        Dim sDetached : sDetached = oSignature.Detached
        Dim sDetachedValue        
        if sDetached then
            sDetachedValue = "������������"
        else sDetachedValue = "�������������"
        end if
        SignatureInformation = SignatureInformation + "��� ���: " + sDetachedValue + vbCrLf
        '����� ������ ��������� ������ CMS
        Dim lCMSVersion : lCMSVersion = oSignature.CMSVersion
        SignatureInformation = SignatureInformation + "����� ������ ��������� CMS: " + CStr(lCMSVersion)
    Next

    '������� ����������
    Set oSignature = Nothing
    Set oSignatures = Nothing

End Function

Function CertificateInformation (ByVal InputFileNameSignature)
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim n : n = oSignatures.Count
    For i=0 to n-1
        Set oSignature = oSignatures.Item(i)

        '�������� ��������� �� ���������� �������
        Set oCertificate = oSignature.Certificate

        '�������� �����������
        CertificateInformation =                          "�������� ����� �����������: " + CStr(oCertificate.SerialNumber) + vbCrLf + vbCrLf
        CertificateInformation = CertificateInformation + "����� ��: " + CStr(oCertificate.IssuerName) + vbCrLf + vbCrLf
        CertificateInformation = CertificateInformation + "�������� �����������: " + CStr(oCertificate.SubjectName) + vbCrLf + vbCrLf
        CertificateInformation = CertificateInformation + "��������� �: " + CStr(oCertificate.ValidFrom) + vbCrLf + vbCrLf
        CertificateInformation = CertificateInformation + "��������� ��: " + CStr(oCertificate.ValidTo)
    Next

    '������� ����������
    Set oCertificate = Nothing
    Set oSignature = Nothing
    Set oSignatures = Nothing

End Function

Function CertificateVerify (ByVal InputFileNameSignature)
    oPKCS7Message.Load DT_SIGNED_DATA, InputFileNameSignature, ""
    Set oSignatures = oPKCS7Message.Signatures
    Dim n : n = oSignatures.Count
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
                CertificateVerify = "������ �����������: " + "���������"
            Case VS_UNSUFFICIENT_INFO
                CertificateVerify = "������ �����������: " + "������ ����������"
            Case VS_UNCORRECT
                CertificateVerify = "������ �����������: " + "�����������"
            Case VS_INVALID_CERTIFICATE_BLOB
                CertificateVerify = "������ �����������: " + "���������������� ���� �����������"
            Case VS_CERTIFICATE_TIME_EXPIRIED
                CertificateVerify = "������ �����������: " + "����� �������� ����������� ������� ��� ��� �� ���������"
            Case VS_CERTIFICATE_NO_CHAIN
                CertificateVerify = "������ �����������: " + "���������� ��������� ������� ������������"
            Case VS_CERTIFICATE_CRL_UPDATING_ERROR
                CertificateVerify = "������ �����������: " + "������ ���������� �����������"
            Case VS_LOCAL_CRL_NOT_FOUND
                CertificateVerify = "������ �����������: " + "�� ������ ��������� ���"
            Case VS_CRL_TIME_EXPIRIED
                CertificateVerify = "������ �����������: " + "������� ����� �������� ���"
            Case VS_CERTIFICATE_IN_CRL
                CertificateVerify = "������ �����������: " + "���������� ��������� � ���"
            Case VS_CERTIFICATE_IN_LOCAL_CRL
                CertificateVerify = "������ �����������: " + "���������� ��������� � ��������� ���"
            Case VS_CERTIFICATE_CORRECT_BY_LOCAL_CRL
                CertificateVerify = "������ �����������: " + "���������� ������������ �� ���������� ���"
            Case VS_CERTIFICATE_USING_RESTRICTED
                CertificateVerify = "������ �����������: " + "������������� ����������� ����������"
        End Select
    Next

    '������� ����������
    Set oCertificate = Nothing
    Set oSignature = Nothing
    Set oSignatures = Nothing

End Function