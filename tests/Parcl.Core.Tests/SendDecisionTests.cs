using Parcl.Core.Crypto;
using Xunit;

namespace Parcl.Core.Tests
{
    public class SendDecisionTests
    {
        // ── Encryption decision tests ──

        [Fact]
        public void Evaluate_ParclEncryptFlag_ShouldEncrypt()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: false);

            Assert.True(decision.ShouldEncrypt);
            Assert.Equal("user-toggle", decision.EncryptSource);
        }

        [Fact]
        public void Evaluate_AlwaysEncrypt_ShouldEncrypt()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: true, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: false);

            Assert.True(decision.ShouldEncrypt);
            Assert.Equal("always-encrypt", decision.EncryptSource);
        }

        [Fact]
        public void Evaluate_NoFlags_ShouldNotEncrypt()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: false);

            Assert.False(decision.ShouldEncrypt);
            Assert.False(decision.ShouldSign);
            Assert.Null(decision.EncryptSource);
            Assert.Null(decision.SignSource);
        }

        [Fact]
        public void Evaluate_AlwaysEncrypt_OverridesAbsentToggle()
        {
            // COM automation: no ParclEncrypt flag, but AlwaysEncrypt is on
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: true, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.UseNativeSmime);
            Assert.Equal("always-encrypt", decision.EncryptSource);
        }

        [Fact]
        public void Evaluate_ParclEncryptFlag_TakesPriority()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: false,
                alwaysEncrypt: true, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: false);

            Assert.True(decision.ShouldEncrypt);
            Assert.Equal("user-toggle", decision.EncryptSource);
        }

        // ── Signing decision tests ──

        [Fact]
        public void Evaluate_ParclSignFlag_ShouldSign()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: true,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: true, useNativeSmime: false);

            Assert.True(decision.ShouldSign);
            Assert.Equal("user-toggle", decision.SignSource);
        }

        [Fact]
        public void Evaluate_AlwaysSign_WithCert_ShouldSign()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: true,
                hasSigningCert: true, useNativeSmime: false);

            Assert.True(decision.ShouldSign);
            Assert.Equal("always-sign", decision.SignSource);
        }

        [Fact]
        public void Evaluate_AlwaysSign_WithoutCert_ShouldNotSign()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: true,
                hasSigningCert: false, useNativeSmime: false);

            Assert.False(decision.ShouldSign);
            Assert.Null(decision.SignSource);
        }

        // ── Combined tests ──

        [Fact]
        public void Evaluate_EncryptAndSign_BothEnabled()
        {
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: true,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: true, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.True(decision.UseNativeSmime);
        }

        [Fact]
        public void Evaluate_AlwaysEncryptAndAlwaysSign_BothEnabled()
        {
            // Full policy mode — COM automation sends with no flags
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: true, alwaysSign: true,
                hasSigningCert: true, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.Equal("always-encrypt", decision.EncryptSource);
            Assert.Equal("always-sign", decision.SignSource);
        }

        [Fact]
        public void Evaluate_NativeSmime_PreservedInDecision()
        {
            var native = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: true);

            var parcl = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: false);

            Assert.True(native.UseNativeSmime);
            Assert.False(parcl.UseNativeSmime);
        }

        // ── PR_SECURITY_FLAGS calculation tests ──

        [Fact]
        public void CalculateSecurityFlags_Encrypt_SetsFlag01()
        {
            int flags = SendDecision.CalculateSecurityFlags(0, encrypt: true, sign: false);
            Assert.Equal(0x01, flags);
        }

        [Fact]
        public void CalculateSecurityFlags_Sign_SetsFlag02()
        {
            int flags = SendDecision.CalculateSecurityFlags(0, encrypt: false, sign: true);
            Assert.Equal(0x02, flags);
        }

        [Fact]
        public void CalculateSecurityFlags_EncryptAndSign_SetsBothFlags()
        {
            int flags = SendDecision.CalculateSecurityFlags(0, encrypt: true, sign: true);
            Assert.Equal(0x03, flags);
        }

        [Fact]
        public void CalculateSecurityFlags_PreservesExistingFlags()
        {
            int existing = 0x10; // some existing flag
            int flags = SendDecision.CalculateSecurityFlags(existing, encrypt: true, sign: false);
            Assert.Equal(0x11, flags);
        }

        [Fact]
        public void CalculateSecurityFlags_NoFlags_ReturnsExisting()
        {
            int existing = 0x04;
            int flags = SendDecision.CalculateSecurityFlags(existing, encrypt: false, sign: false);
            Assert.Equal(0x04, flags);
        }

        [Fact]
        public void CalculateSecurityFlags_IdempotentWhenAlreadySet()
        {
            int existing = 0x03; // already has encrypt + sign
            int flags = SendDecision.CalculateSecurityFlags(existing, encrypt: true, sign: true);
            Assert.Equal(0x03, flags); // OR is idempotent
        }

        // ── Native S/MIME routing regression tests ──

        [Fact]
        public void Evaluate_NativeSmime_EncryptOnly_ShouldUseNative()
        {
            // Regression: native encrypt-only must NOT fall back to Parcl envelope
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: false,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: false, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.False(decision.ShouldSign);
            Assert.True(decision.UseNativeSmime);
        }

        [Fact]
        public void Evaluate_NativeSmime_SignOnly_ShouldUseNative()
        {
            // Sign-only with native S/MIME should use PR_SECURITY_FLAGS
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: true,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: true, useNativeSmime: true);

            Assert.False(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.True(decision.UseNativeSmime);
        }

        [Fact]
        public void Evaluate_NativeSmime_EncryptAndSign_ShouldUseNative()
        {
            // Encrypt+sign with native S/MIME should use PR_SECURITY_FLAGS for both
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: true,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: true, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.True(decision.UseNativeSmime);
        }

        [Fact]
        public void Evaluate_NonNative_EncryptAndSign_UsesParclEnvelope()
        {
            // Without native S/MIME, encrypt+sign uses Parcl envelope (sign-then-encrypt)
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: true, parclSignFlag: true,
                alwaysEncrypt: false, alwaysSign: false,
                hasSigningCert: true, useNativeSmime: false);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.False(decision.UseNativeSmime);
        }

        [Fact]
        public void Evaluate_AlwaysEncrypt_NativeSmime_ShouldRoute_Native()
        {
            // Policy: always encrypt + native S/MIME → native path
            var decision = SendDecision.Evaluate(
                parclEncryptFlag: false, parclSignFlag: false,
                alwaysEncrypt: true, alwaysSign: true,
                hasSigningCert: true, useNativeSmime: true);

            Assert.True(decision.ShouldEncrypt);
            Assert.True(decision.ShouldSign);
            Assert.True(decision.UseNativeSmime);
            Assert.Equal("always-encrypt", decision.EncryptSource);
            Assert.Equal("always-sign", decision.SignSource);
        }
    }
}
