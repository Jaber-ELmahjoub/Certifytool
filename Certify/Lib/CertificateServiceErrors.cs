using System;
using System.Collections.Generic;

namespace Certify.Lib
{
    internal static class CertificateServiceErrors
    {
        // Based on Microsoft Certificate Services HRESULT documentation:
        // https://learn.microsoft.com/en-us/windows/win32/com/com-error-codes-4
        private static readonly Dictionary<uint, Tuple<string, string>> Errors =
            new Dictionary<uint, Tuple<string, string>>
            {
                { 0x80094005, Tuple.Create("CERTSRV_E_BAD_CERTIFICATE", "The certification authority's certificate contains invalid data.") },
                { 0x80094006, Tuple.Create("CERTSRV_E_SERVER_SUSPENDED", "Certificate service has been suspended for a database restore operation.") },
                { 0x80094007, Tuple.Create("CERTSRV_E_ENCODING_LENGTH", "The certificate contains an encoded length that is potentially incompatible with older enrollment software.") },
                { 0x80094008, Tuple.Create("CERTSRV_E_ROLECONFLICT", "The operation is denied. The user has multiple roles assigned and the certification authority is configured to enforce role separation.") },
                { 0x80094009, Tuple.Create("CERTSRV_E_RESTRICTEDOFFICER", "The operation is denied. It can only be performed by a certificate manager that is allowed to manage certificates for the current requester.") },
                { 0x8009400A, Tuple.Create("CERTSRV_E_KEY_ARCHIVAL_NOT_CONFIGURED", "Cannot archive private key. The certification authority is not configured for key archival.") },
                { 0x8009400B, Tuple.Create("CERTSRV_E_NO_VALID_KRA", "Cannot archive private key. The certification authority could not verify one or more key recovery certificates.") },
                { 0x8009400C, Tuple.Create("CERTSRV_E_BAD_REQUEST_KEY_ARCHIVAL", "The request is incorrectly formatted. The encrypted private key must be in an unauthenticated attribute in an outermost signature.") },
                { 0x8009400D, Tuple.Create("CERTSRV_E_NO_CAADMIN_DEFINED", "At least one security principal must have the permission to manage this CA.") },
                { 0x8009400E, Tuple.Create("CERTSRV_E_BAD_RENEWAL_CERT_ATTRIBUTE", "The request contains an invalid renewal certificate attribute.") },
                { 0x8009400F, Tuple.Create("CERTSRV_E_NO_DB_SESSIONS", "An attempt was made to open a Certification Authority database session, but there are already too many active sessions.") },
                { 0x80094010, Tuple.Create("CERTSRV_E_ALIGNMENT_FAULT", "A memory reference caused a data alignment fault.") },
                { 0x80094011, Tuple.Create("CERTSRV_E_ENROLL_DENIED", "The permissions on this certification authority do not allow the current user to enroll for certificates.") },
                { 0x80094012, Tuple.Create("CERTSRV_E_TEMPLATE_DENIED", "The permissions on the certificate template do not allow the current user to enroll for this type of certificate.") },
                { 0x80094013, Tuple.Create("CERTSRV_E_DOWNLEVEL_DC_SSL_OR_UPGRADE", "The contacted domain controller cannot support signed LDAP traffic. Update the domain controller or configure Certificate Services to use SSL for Active Directory access.") },
                { 0x80094014, Tuple.Create("CERTSRV_E_ADMIN_DENIED_REQUEST", "The request was denied by a certificate manager or CA administrator.") },
                { 0x80094015, Tuple.Create("CERTSRV_E_NO_POLICY_SERVER", "An enrollment policy server cannot be located.") },
                { 0x80094800, Tuple.Create("CERTSRV_E_UNSUPPORTED_CERT_TYPE", "The requested certificate template is not supported by this CA.") },
                { 0x80094801, Tuple.Create("CERTSRV_E_NO_CERT_TYPE", "The request contains no certificate template information.") },
                { 0x80094802, Tuple.Create("CERTSRV_E_TEMPLATE_CONFLICT", "The request contains conflicting certificate template information.") },
                { 0x80094803, Tuple.Create("CERTSRV_E_SUBJECT_ALT_NAME_REQUIRED", "The request is missing a required Subject Alternate name extension.") },
                { 0x80094804, Tuple.Create("CERTSRV_E_ARCHIVED_KEY_REQUIRED", "The request is missing a required private key for archival by the server.") },
                { 0x80094805, Tuple.Create("CERTSRV_E_SMIME_REQUIRED", "The request is missing a required SMIME capabilities extension.") },
                { 0x80094806, Tuple.Create("CERTSRV_E_BAD_RENEWAL_SUBJECT", "The request was made on behalf of a subject other than the caller. The certificate template must be configured to require at least one signature to authorize the request.") },
                { 0x80094807, Tuple.Create("CERTSRV_E_BAD_TEMPLATE_VERSION", "The request template version is newer than the supported template version.") },
                { 0x80094808, Tuple.Create("CERTSRV_E_TEMPLATE_POLICY_REQUIRED", "The template is missing a required signature policy attribute.") },
                { 0x80094809, Tuple.Create("CERTSRV_E_SIGNATURE_POLICY_REQUIRED", "A signature is required because the template's issuance policies demand an authorized signature.") },
                { 0x8009480A, Tuple.Create("CERTSRV_E_SIGNATURE_COUNT", "The request is missing one or more required signatures.") },
                { 0x8009480B, Tuple.Create("CERTSRV_E_SIGNATURE_REJECTED", "One or more signatures did not include the required application or issuance policies. The request is missing one or more required valid signatures.") },
                { 0x8009480C, Tuple.Create("CERTSRV_E_ISSUANCE_POLICY_REQUIRED", "The request is missing one or more required signature issuance policies.") },
                { 0x8009480D, Tuple.Create("CERTSRV_E_SUBJECT_UPN_REQUIRED", "A UPN value is required and could not be added to the subject alternative name.") },
                { 0x8009480E, Tuple.Create("CERTSRV_E_SUBJECT_DIRECTORY_GUID_REQUIRED", "An Active Directory GUID value is required and could not be added to the subject alternative name.") },
                { 0x8009480F, Tuple.Create("CERTSRV_E_SUBJECT_DNS_REQUIRED", "A DNS name is required and could not be added to the subject alternative name.") },
                { 0x80094810, Tuple.Create("CERTSRV_E_ARCHIVED_KEY_UNEXPECTED", "The request includes a private key for archival, but key archival is not enabled for the specified certificate template.") },
                { 0x80094811, Tuple.Create("CERTSRV_E_KEY_LENGTH", "The public key does not meet the minimum size required by the specified certificate template.") },
                { 0x80094812, Tuple.Create("CERTSRV_E_SUBJECT_EMAIL_REQUIRED", "An email name is required and could not be added to the subject or subject alternative name.") },
                { 0x80094813, Tuple.Create("CERTSRV_E_UNKNOWN_CERT_TYPE", "One or more certificate templates to be enabled on this CA could not be found.") },
                { 0x80094814, Tuple.Create("CERTSRV_E_CERT_TYPE_OVERLAP", "The certificate template renewal period is longer than the certificate validity period. The template should be reconfigured or the CA certificate renewed.") },
                { 0x80094815, Tuple.Create("CERTSRV_E_TOO_MANY_SIGNATURES", "The certificate template requires too many RA signatures. Only one RA signature is allowed.") },
                { 0x80094816, Tuple.Create("CERTSRV_E_RENEWAL_BAD_PUBLIC_KEY", "The certificate template requires renewal with the same public key, but the request uses a different public key.") }
            };

        public static bool TryGetError(uint hresult, out string errorName, out string errorMessage)
        {
            if (Errors.TryGetValue(hresult, out var error))
            {
                errorName = error.Item1;
                errorMessage = error.Item2;
                return true;
            }

            errorName = null;
            errorMessage = null;
            return false;
        }
    }
}
