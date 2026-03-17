namespace Parcl.Core.Crypto
{
    /// <summary>
    /// Determines whether a message should be encrypted and/or signed at send time.
    /// Extracted from the add-in's ItemSend handler to enable unit testing without Outlook COM.
    /// </summary>
    public class SendDecision
    {
        public bool ShouldEncrypt { get; set; }
        public bool ShouldSign { get; set; }
        public bool UseNativeSmime { get; set; }
        public string? EncryptSource { get; set; }
        public string? SignSource { get; set; }

        /// <summary>
        /// Evaluates all encryption/signing inputs and returns a decision.
        /// </summary>
        /// <param name="parclEncryptFlag">True if the ParclEncrypt UserProperty is set on the mail item (ribbon toggle).</param>
        /// <param name="parclSignFlag">True if the ParclSign UserProperty is set on the mail item (ribbon toggle).</param>
        /// <param name="alwaysEncrypt">True if the AlwaysEncrypt setting is enabled (persistent policy).</param>
        /// <param name="alwaysSign">True if the AlwaysSign setting is enabled (persistent policy).</param>
        /// <param name="hasSigningCert">True if a signing certificate is configured and has a private key.</param>
        /// <param name="useNativeSmime">True if native Outlook S/MIME mode is selected.</param>
        public static SendDecision Evaluate(
            bool parclEncryptFlag,
            bool parclSignFlag,
            bool alwaysEncrypt,
            bool alwaysSign,
            bool hasSigningCert,
            bool useNativeSmime)
        {
            var decision = new SendDecision { UseNativeSmime = useNativeSmime };

            // Encryption: any of these sources triggers encryption
            if (parclEncryptFlag)
            {
                decision.ShouldEncrypt = true;
                decision.EncryptSource = "user-toggle";
            }
            else if (alwaysEncrypt)
            {
                decision.ShouldEncrypt = true;
                decision.EncryptSource = "always-encrypt";
            }

            // Signing: ribbon toggle OR AlwaysSign with a valid cert
            if (parclSignFlag)
            {
                decision.ShouldSign = true;
                decision.SignSource = "user-toggle";
            }
            else if (alwaysSign && hasSigningCert)
            {
                decision.ShouldSign = true;
                decision.SignSource = "always-sign";
            }

            return decision;
        }

        /// <summary>
        /// Calculates PR_SECURITY_FLAGS value for native S/MIME mode.
        /// </summary>
        /// <param name="existingFlags">Current PR_SECURITY_FLAGS value on the mail item.</param>
        /// <param name="encrypt">Whether to add encryption flag.</param>
        /// <param name="sign">Whether to add signing flag.</param>
        /// <returns>Updated flags value with encryption and/or signing bits set.</returns>
        public static int CalculateSecurityFlags(int existingFlags, bool encrypt, bool sign)
        {
            int flags = existingFlags;
            if (encrypt) flags |= 0x01; // SECFLAG_ENCRYPTED
            if (sign) flags |= 0x02;    // SECFLAG_SIGNED
            return flags;
        }
    }
}
