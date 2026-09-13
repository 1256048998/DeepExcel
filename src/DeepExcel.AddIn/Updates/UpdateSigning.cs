using System;
using System.Security.Cryptography;
using System.Text;

namespace DeepExcel.AddIn.Updates
{
    /// <summary>Verifies the detached signature over an update manifest payload.</summary>
    public interface IManifestVerifier
    {
        /// <summary>False when no public key was compiled in; the updater then refuses everything.</summary>
        bool IsConfigured { get; }

        /// <summary>Derived from the key material, so it can never drift from the key it names.</summary>
        string KeyId { get; }

        bool Verify(byte[] signedBytes, byte[] signature);
    }

    /// <summary>
    /// RSA-PSS / SHA-256 verification over raw public key parameters.
    ///
    /// **Why RSA rather than the Ed25519 named in the roadmap.** .NET Framework
    /// 4.8 has no Ed25519 primitive, so it would mean either hand-rolled curve
    /// arithmetic (unacceptable for signature verification) or shipping
    /// BouncyCastle — a multi-megabyte third-party DLL added to a payload we
    /// already know draws antivirus false positives, which is the very friction
    /// the auto-updater exists to remove. RSA-4096/PSS/SHA-256 runs on Windows
    /// CNG, needs no extra file, and is equally sufficient against the actual
    /// threat here: a poisoned update channel.
    ///
    /// The key is carried as modulus + exponent rather than a
    /// SubjectPublicKeyInfo blob because .NET Framework cannot import SPKI
    /// without an ASN.1 parser, and a hand-written DER parser sitting in front
    /// of a trust decision is exactly the wrong place to save a dependency.
    /// </summary>
    public sealed class RsaManifestVerifier : IManifestVerifier
    {
        private readonly byte[] _modulus;
        private readonly byte[] _exponent;

        private RsaManifestVerifier(byte[] modulus, byte[] exponent, string keyId)
        {
            _modulus = modulus;
            _exponent = exponent;
            KeyId = keyId;
        }

        public bool IsConfigured => _modulus != null;

        public string KeyId { get; }

        /// <summary>
        /// Builds a verifier, or an unconfigured one when the key is absent or
        /// unusable. Never throws: a malformed compiled-in key must disable
        /// updates, not crash the add-in on startup.
        /// </summary>
        public static RsaManifestVerifier FromBase64(string modulusBase64, string exponentBase64)
        {
            if (string.IsNullOrWhiteSpace(modulusBase64) || string.IsNullOrWhiteSpace(exponentBase64))
            {
                return new RsaManifestVerifier(null, null, null);
            }

            byte[] modulus, exponent;
            try
            {
                modulus = Convert.FromBase64String(modulusBase64.Trim());
                exponent = Convert.FromBase64String(exponentBase64.Trim());
            }
            catch (FormatException)
            {
                return new RsaManifestVerifier(null, null, null);
            }

            // 2048-bit floor; anything shorter is not worth the code that reads it.
            if (modulus.Length < 256 || modulus.Length > 1024 || exponent.Length == 0 || exponent.Length > 8)
            {
                return new RsaManifestVerifier(null, null, null);
            }

            return new RsaManifestVerifier(modulus, exponent, ComputeKeyId(modulus, exponent));
        }

        /// <summary>
        /// First 8 bytes of SHA-256(modulus || exponent), lowercase hex.
        ///
        /// Both sides compute it from the key bytes, so a manifest signed with
        /// the wrong private key is rejected by key id before any signature
        /// math runs — and the release tooling can catch the mismatch at build
        /// time instead of on user machines.
        /// </summary>
        public static string ComputeKeyId(byte[] modulus, byte[] exponent)
        {
            using (var sha = SHA256.Create())
            {
                sha.TransformBlock(modulus, 0, modulus.Length, null, 0);
                sha.TransformFinalBlock(exponent, 0, exponent.Length);
                byte[] hash = sha.Hash;
                var builder = new StringBuilder(16);
                for (int i = 0; i < 8; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }
                return builder.ToString();
            }
        }

        public bool Verify(byte[] signedBytes, byte[] signature)
        {
            if (!IsConfigured || signedBytes == null || signature == null || signature.Length == 0)
            {
                return false;
            }

            try
            {
                // A fresh RSACng per call: verification happens about once per
                // session, and a per-call instance removes every question about
                // handle lifetime and thread affinity.
                using (var rsa = new RSACng())
                {
                    rsa.ImportParameters(new RSAParameters { Modulus = _modulus, Exponent = _exponent });
                    // Salt length must match the hash length on both sides.
                    // .NET fixes PSS salt length at the digest size; the Python
                    // signer must therefore use PSS.DIGEST_LENGTH and not
                    // cryptography's MAX_LENGTH, or every signature fails here
                    // with no explanation. Locked down by UpdateVectorTests,
                    // which verifies manifests produced by the real signer.
                    return rsa.VerifyData(
                        signedBytes, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pss);
                }
            }
            catch (CryptographicException)
            {
                return false;
            }
            catch (PlatformNotSupportedException)
            {
                return false;
            }
        }
    }

    /// <summary>
    /// The update-signing public key compiled into the client.
    ///
    /// Empty by default, which disables updating entirely — a build that has
    /// never been given a key must not accept an unsigned or foreign manifest.
    /// Populate it with:
    ///
    ///     python scripts/update_signing.py genkey --out &lt;offline path&gt;
    ///
    /// which rewrites the two constants below in place so the compiled-in public
    /// key and the private key that signs releases cannot be edited apart by
    /// hand. The private key never belongs in this repository, on the server, or
    /// on a build machine that is reachable from the internet: the server only
    /// relays an already-signed manifest, so compromising the server does not
    /// let anyone push an update.
    /// </summary>
    public static class EmbeddedUpdateKey
    {
        // BEGIN GENERATED UPDATE KEY -- scripts/update_signing.py rewrites these two lines
        public const string ModulusBase64 = "";
        public const string ExponentBase64 = "";
        // END GENERATED UPDATE KEY

        private static readonly RsaManifestVerifier Instance =
            RsaManifestVerifier.FromBase64(ModulusBase64, ExponentBase64);

        public static IManifestVerifier Verifier => Instance;
    }
}
