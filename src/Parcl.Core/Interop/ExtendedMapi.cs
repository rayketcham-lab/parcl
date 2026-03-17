using System;
using System.Runtime.InteropServices;
using Parcl.Core.Config;

namespace Parcl.Core.Interop
{
    /// <summary>
    /// Extended MAPI P/Invoke wrapper for setting raw S/MIME content on an
    /// Outlook MailItem via its underlying IMessage COM object.
    ///
    /// Outlook stores S/MIME messages as hidden attachments on the MAPI message:
    ///   - PR_MESSAGE_CLASS = "IPM.Note.SMIME"
    ///   - A hidden attachment (PR_RENDERING_POSITION = -1) containing the CMS
    ///     envelope bytes with PR_ATTACH_MIME_TAG = "application/pkcs7-mime"
    ///
    /// This allows Outlook to render the decrypted content inline in the reading
    /// pane instead of presenting a raw smime.p7m file attachment.
    ///
    /// Usage:
    ///   object mapiObject = mailItem.MAPIOBJECT;
    ///   ExtendedMapi.SetSmimeContent(mapiObject, cmsEnvelopeBytes, logger);
    /// </summary>
    public static class ExtendedMapi
    {
        private const string LogComponent = "MAPI";

        // ── MAPI property tags ──────────────────────────────────────────
        private const uint PR_MESSAGE_CLASS_W     = 0x001A001F; // PT_UNICODE
        private const uint PR_BODY_W              = 0x1000001F; // PT_UNICODE
        private const uint PR_BODY_HTML           = 0x10130102; // PT_BINARY
        private const uint PR_RTF_COMPRESSED      = 0x10090102; // PT_BINARY
        private const uint PR_ATTACH_METHOD       = 0x37050003; // PT_LONG
        private const uint PR_ATTACH_FILENAME_W   = 0x3704001F; // PT_UNICODE
        private const uint PR_ATTACH_LONG_FILENAME_W = 0x3707001F; // PT_UNICODE
        private const uint PR_ATTACH_MIME_TAG_W   = 0x370E001F; // PT_UNICODE
        private const uint PR_ATTACH_DATA_BIN     = 0x37010102; // PT_BINARY
        private const uint PR_RENDERING_POSITION  = 0x370B0003; // PT_LONG

        // MAPI attach methods
        private const int ATTACH_BY_VALUE = 1;

        // MAPI flags
        private const int MAPI_BEST_ACCESS = 0x00000010;
        private const int MAPI_MODIFY      = 0x00000001;

        // S_OK
        private const int S_OK = 0;

        // IMessage GUID: {00020307-0000-0000-C000-000000000046}
        private static readonly Guid IID_IMessage =
            new Guid("00020307-0000-0000-C000-000000000046");

        // ── IMAPIProp vtable offsets (IUnknown=0-2, then IMAPIProp methods) ─
        //  0: QueryInterface  1: AddRef       2: Release
        //  3: GetLastError    4: SaveChanges  5: GetProps
        //  6: GetPropList     7: OpenProperty 8: SetProps
        //  9: DeleteProps    10: CopyTo      11: CopyProps
        // 12: GetNamesFromIDs 13: GetIDsFromNames
        //
        // IMessage extends IMAPIProp, additional methods start at index 14:
        // 14: GetAttachmentTable  15: OpenAttach  16: CreateAttach
        // 17: DeleteAttach        18: SubmitMessage
        private const int VtableIndex_SaveChanges          = 4;
        private const int VtableIndex_SetProps             = 8;
        private const int VtableIndex_DeleteProps          = 9;
        private const int VtableIndex_CreateAttach         = 16;

        // ── SPropValue layout (matches 64-bit MAPI) ────────────────────
        //
        // typedef struct {
        //     ULONG ulPropTag;   // 4 bytes
        //     ULONG dwAlignPad;  // 4 bytes
        //     union { ... }      // 8 bytes on x64
        // } SPropValue;

        [StructLayout(LayoutKind.Sequential)]
        private struct SPropValue
        {
            public uint ulPropTag;
            public uint dwAlignPad;
            public SPropValueData Value;
        }

        [StructLayout(LayoutKind.Explicit, Size = 8)]
        private struct SPropValueData
        {
            [FieldOffset(0)] public int l;       // PT_LONG
            [FieldOffset(0)] public IntPtr lpszW; // PT_UNICODE (pointer to wchar_t*)
            [FieldOffset(0)] public long li;     // PT_I8 / alignment
        }

        // SBinary struct for PT_BINARY properties
        [StructLayout(LayoutKind.Sequential)]
        private struct SBinary
        {
            public uint cb;
            public IntPtr lpb;
        }

        // SPropTagArray for DeleteProps
        [StructLayout(LayoutKind.Sequential)]
        private struct SPropTagArray
        {
            public uint cValues;
            // Followed by uint[] aulPropTag — we marshal manually
        }

        // ── P/Invoke: mapi32.dll / olmapi32.dll ────────────────────────

        [DllImport("mapi32.dll", EntryPoint = "HrSetOneProp@8", CallingConvention = CallingConvention.StdCall)]
        private static extern int HrSetOneProp_Mapi32(IntPtr lpMapiProp, ref SPropValue lpPropValue);

        [DllImport("olmapi32.dll", EntryPoint = "HrSetOneProp", CallingConvention = CallingConvention.StdCall)]
        private static extern int HrSetOneProp_OlMapi32(IntPtr lpMapiProp, ref SPropValue lpPropValue);

        private static bool _useOlMapi32;
        private static bool _mapiResolved;

        /// <summary>
        /// Calls HrSetOneProp, trying mapi32.dll first, then olmapi32.dll.
        /// </summary>
        private static int HrSetOneProp(IntPtr lpMapiProp, ref SPropValue prop)
        {
            if (!_mapiResolved)
            {
                // Outlook installs olmapi32.dll; standalone MAPI uses mapi32.dll.
                // Try the Outlook-specific DLL first since this is a VSTO add-in.
                try
                {
                    int hr = HrSetOneProp_OlMapi32(lpMapiProp, ref prop);
                    _useOlMapi32 = true;
                    _mapiResolved = true;
                    return hr;
                }
                catch (DllNotFoundException)
                {
                    _useOlMapi32 = false;
                }
                catch (EntryPointNotFoundException)
                {
                    _useOlMapi32 = false;
                }

                _mapiResolved = true;
                return HrSetOneProp_Mapi32(lpMapiProp, ref prop);
            }

            return _useOlMapi32
                ? HrSetOneProp_OlMapi32(lpMapiProp, ref prop)
                : HrSetOneProp_Mapi32(lpMapiProp, ref prop);
        }

        // ── Vtable call delegates ──────────────────────────────────────
        // These match the COM method signatures we need to call through
        // the IMessage vtable.

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int SaveChangesDelegate(IntPtr pThis, uint ulFlags);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int SetPropsDelegate(
            IntPtr pThis, uint cValues, IntPtr lpPropArray, IntPtr lppProblems);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int DeletePropsDelegate(
            IntPtr pThis, IntPtr lpPropTagArray, IntPtr lppProblems);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int CreateAttachDelegate(
            IntPtr pThis, IntPtr lpInterface, uint ulFlags,
            out uint lpulAttachmentNum, out IntPtr lppAttach);

        // ── Public API ─────────────────────────────────────────────────

        /// <summary>
        /// Converts an Outlook MailItem's underlying MAPI message into a proper
        /// S/MIME message by setting PR_MESSAGE_CLASS and creating a hidden
        /// attachment containing the CMS envelope bytes.
        ///
        /// This causes Outlook to recognize the message as S/MIME and render the
        /// decrypted content inline rather than showing a smime.p7m attachment.
        /// </summary>
        /// <param name="mapiObject">
        /// The MailItem.MAPIOBJECT property value (IUnknown pointer to IMessage).
        /// </param>
        /// <param name="cmsEnvelope">
        /// DER-encoded CMS EnvelopedData bytes (the S/MIME encrypted payload).
        /// </param>
        /// <param name="logger">Optional logger instance.</param>
        /// <returns>True if the MAPI properties were set successfully.</returns>
        public static bool SetSmimeContent(
            object mapiObject, byte[] cmsEnvelope, ParclLogger? logger = null)
        {
            if (mapiObject == null)
                throw new ArgumentNullException(nameof(mapiObject));
            if (cmsEnvelope == null || cmsEnvelope.Length == 0)
                throw new ArgumentException("CMS envelope data must not be empty.", nameof(cmsEnvelope));

            IntPtr pUnk = IntPtr.Zero;
            IntPtr pMsg = IntPtr.Zero;

            try
            {
                pUnk = Marshal.GetIUnknownForObject(mapiObject);
                var iid = IID_IMessage;
                int hr = Marshal.QueryInterface(pUnk, ref iid, out pMsg);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"QueryInterface for IMessage failed: 0x{hr:X8}");
                    return false;
                }

                logger?.Debug(LogComponent, "Acquired IMessage interface from MAPIOBJECT");

                // Step 1: Delete existing body properties so Outlook does not
                //         display stale plain-text or HTML alongside the S/MIME blob.
                if (!DeleteBodyProperties(pMsg, logger))
                {
                    logger?.Warn(LogComponent,
                        "DeleteProps for body properties returned an error — continuing");
                }

                // Step 2: Set PR_MESSAGE_CLASS to "IPM.Note.SMIME"
                if (!SetStringProperty(pMsg, PR_MESSAGE_CLASS_W, "IPM.Note.SMIME", logger))
                {
                    logger?.Error(LogComponent, "Failed to set PR_MESSAGE_CLASS");
                    return false;
                }

                logger?.Debug(LogComponent, "PR_MESSAGE_CLASS set to IPM.Note.SMIME");

                // Step 3: Create hidden attachment with CMS envelope
                if (!CreateSmimeAttachment(pMsg, cmsEnvelope, logger))
                {
                    logger?.Error(LogComponent, "Failed to create S/MIME attachment");
                    return false;
                }

                // Step 4: SaveChanges on IMessage
                if (!CallSaveChanges(pMsg, logger))
                {
                    logger?.Error(LogComponent, "IMessage::SaveChanges failed");
                    return false;
                }

                logger?.Info(LogComponent,
                    $"S/MIME content set successfully ({cmsEnvelope.Length} bytes)");
                return true;
            }
            catch (COMException ex)
            {
                logger?.Error(LogComponent, "COM error in SetSmimeContent", ex);
                return false;
            }
            catch (Exception ex)
            {
                logger?.Error(LogComponent, "Unexpected error in SetSmimeContent", ex);
                return false;
            }
            finally
            {
                if (pMsg != IntPtr.Zero) Marshal.Release(pMsg);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        /// <summary>
        /// Sets a single MAPI string (PT_UNICODE) property via HrSetOneProp.
        /// </summary>
        /// <param name="mapiObject">
        /// The MailItem.MAPIOBJECT property value.
        /// </param>
        /// <param name="propTag">MAPI property tag (must be PT_UNICODE / 0x001F type).</param>
        /// <param name="value">The string value to set.</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <returns>True on success.</returns>
        public static bool SetStringProperty(
            object mapiObject, uint propTag, string value, ParclLogger? logger = null)
        {
            if (mapiObject == null)
                throw new ArgumentNullException(nameof(mapiObject));

            IntPtr pUnk = IntPtr.Zero;
            IntPtr pMsg = IntPtr.Zero;

            try
            {
                pUnk = Marshal.GetIUnknownForObject(mapiObject);
                var iid = IID_IMessage;
                int hr = Marshal.QueryInterface(pUnk, ref iid, out pMsg);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"QueryInterface for IMessage failed: 0x{hr:X8}");
                    return false;
                }

                return SetStringProperty(pMsg, propTag, value, logger);
            }
            finally
            {
                if (pMsg != IntPtr.Zero) Marshal.Release(pMsg);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        /// <summary>
        /// Sets a single MAPI integer (PT_LONG) property via HrSetOneProp.
        /// </summary>
        /// <param name="mapiObject">
        /// The MailItem.MAPIOBJECT property value.
        /// </param>
        /// <param name="propTag">MAPI property tag (must be PT_LONG / 0x0003 type).</param>
        /// <param name="value">The integer value to set.</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <returns>True on success.</returns>
        public static bool SetIntProperty(
            object mapiObject, uint propTag, int value, ParclLogger? logger = null)
        {
            if (mapiObject == null)
                throw new ArgumentNullException(nameof(mapiObject));

            IntPtr pUnk = IntPtr.Zero;
            IntPtr pMsg = IntPtr.Zero;

            try
            {
                pUnk = Marshal.GetIUnknownForObject(mapiObject);
                var iid = IID_IMessage;
                int hr = Marshal.QueryInterface(pUnk, ref iid, out pMsg);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"QueryInterface for IMessage failed: 0x{hr:X8}");
                    return false;
                }

                return SetIntProperty(pMsg, propTag, value, logger);
            }
            finally
            {
                if (pMsg != IntPtr.Zero) Marshal.Release(pMsg);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        /// <summary>
        /// Sets a single MAPI binary (PT_BINARY) property via the IMessage
        /// vtable SetProps method.
        /// </summary>
        /// <param name="mapiObject">
        /// The MailItem.MAPIOBJECT property value.
        /// </param>
        /// <param name="propTag">MAPI property tag (must be PT_BINARY / 0x0102 type).</param>
        /// <param name="data">The binary data to set.</param>
        /// <param name="logger">Optional logger instance.</param>
        /// <returns>True on success.</returns>
        public static bool SetBinaryProperty(
            object mapiObject, uint propTag, byte[] data, ParclLogger? logger = null)
        {
            if (mapiObject == null)
                throw new ArgumentNullException(nameof(mapiObject));
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            IntPtr pUnk = IntPtr.Zero;
            IntPtr pMsg = IntPtr.Zero;

            try
            {
                pUnk = Marshal.GetIUnknownForObject(mapiObject);
                var iid = IID_IMessage;
                int hr = Marshal.QueryInterface(pUnk, ref iid, out pMsg);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"QueryInterface for IMessage failed: 0x{hr:X8}");
                    return false;
                }

                return SetBinaryProperty(pMsg, propTag, data, logger);
            }
            finally
            {
                if (pMsg != IntPtr.Zero) Marshal.Release(pMsg);
                if (pUnk != IntPtr.Zero) Marshal.Release(pUnk);
            }
        }

        // ── Internal helpers (operate on raw IMessage IntPtr) ──────────

        private static bool SetStringProperty(
            IntPtr pMsg, uint propTag, string value, ParclLogger? logger)
        {
            IntPtr pStr = Marshal.StringToCoTaskMemUni(value);
            try
            {
                var prop = new SPropValue
                {
                    ulPropTag = propTag,
                    dwAlignPad = 0
                };
                prop.Value.lpszW = pStr;

                int hr = HrSetOneProp(pMsg, ref prop);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"HrSetOneProp(0x{propTag:X8}) failed: 0x{hr:X8}");
                    return false;
                }

                return true;
            }
            finally
            {
                Marshal.FreeCoTaskMem(pStr);
            }
        }

        private static bool SetIntProperty(
            IntPtr pMsg, uint propTag, int value, ParclLogger? logger)
        {
            var prop = new SPropValue
            {
                ulPropTag = propTag,
                dwAlignPad = 0
            };
            prop.Value.l = value;

            int hr = HrSetOneProp(pMsg, ref prop);
            if (hr != S_OK)
            {
                logger?.Error(LogComponent,
                    $"HrSetOneProp(0x{propTag:X8}) failed: 0x{hr:X8}");
                return false;
            }

            return true;
        }

        private static bool SetBinaryProperty(
            IntPtr pMsg, uint propTag, byte[] data, ParclLogger? logger)
        {
            // For PT_BINARY, SPropValue.Value is an SBinary { cb, lpb }.
            // We need to marshal the SBinary struct and the byte array into
            // unmanaged memory, then call SetProps via the vtable because
            // HrSetOneProp does not handle binary properties reliably with
            // the SPropValue union on 64-bit.

            IntPtr pData = IntPtr.Zero;
            IntPtr pPropArray = IntPtr.Zero;

            try
            {
                // Allocate and copy binary data
                pData = Marshal.AllocCoTaskMem(data.Length);
                Marshal.Copy(data, 0, pData, data.Length);

                // Build SPropValue with embedded SBinary.
                // Layout: ulPropTag(4) + dwAlignPad(4) + cb(4) + padding(4) + lpb(8)
                // Total: 24 bytes on x64
                int propSize = IntPtr.Size == 8 ? 24 : 16;
                pPropArray = Marshal.AllocCoTaskMem(propSize);

                // Zero the memory to avoid garbage in padding
                for (int i = 0; i < propSize; i++)
                    Marshal.WriteByte(pPropArray, i, 0);

                // ulPropTag at offset 0
                Marshal.WriteInt32(pPropArray, 0, unchecked((int)propTag));
                // dwAlignPad at offset 4 (already zeroed)

                if (IntPtr.Size == 8)
                {
                    // x64: SBinary.cb at offset 8, SBinary.lpb at offset 16
                    Marshal.WriteInt32(pPropArray, 8, data.Length);
                    Marshal.WriteIntPtr(pPropArray, 16, pData);
                }
                else
                {
                    // x86: SBinary.cb at offset 8, SBinary.lpb at offset 12
                    Marshal.WriteInt32(pPropArray, 8, data.Length);
                    Marshal.WriteIntPtr(pPropArray, 12, pData);
                }

                // Call IMessage::SetProps(1, pPropArray, NULL) via vtable
                IntPtr vtable = Marshal.ReadIntPtr(pMsg);
                IntPtr pSetProps = Marshal.ReadIntPtr(vtable,
                    VtableIndex_SetProps * IntPtr.Size);
                var setProps = (SetPropsDelegate)Marshal.GetDelegateForFunctionPointer(
                    pSetProps, typeof(SetPropsDelegate));

                int hr = setProps(pMsg, 1, pPropArray, IntPtr.Zero);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"IMessage::SetProps(0x{propTag:X8}) failed: 0x{hr:X8}");
                    return false;
                }

                return true;
            }
            finally
            {
                if (pData != IntPtr.Zero) Marshal.FreeCoTaskMem(pData);
                if (pPropArray != IntPtr.Zero) Marshal.FreeCoTaskMem(pPropArray);
            }
        }

        /// <summary>
        /// Deletes PR_BODY, PR_BODY_HTML, and PR_RTF_COMPRESSED from the
        /// message so Outlook does not render stale body content alongside
        /// the S/MIME attachment.
        /// </summary>
        private static bool DeleteBodyProperties(IntPtr pMsg, ParclLogger? logger)
        {
            // SPropTagArray: { cValues, aulPropTag[3] }
            // 4 bytes count + 3 * 4 bytes tags = 16 bytes
            IntPtr pTagArray = IntPtr.Zero;

            try
            {
                pTagArray = Marshal.AllocCoTaskMem(16);
                Marshal.WriteInt32(pTagArray, 0, 3); // cValues
                Marshal.WriteInt32(pTagArray, 4, unchecked((int)PR_BODY_W));
                Marshal.WriteInt32(pTagArray, 8, unchecked((int)PR_BODY_HTML));
                Marshal.WriteInt32(pTagArray, 12, unchecked((int)PR_RTF_COMPRESSED));

                IntPtr vtable = Marshal.ReadIntPtr(pMsg);
                IntPtr pDeleteProps = Marshal.ReadIntPtr(vtable,
                    VtableIndex_DeleteProps * IntPtr.Size);
                var deleteProps = (DeletePropsDelegate)Marshal.GetDelegateForFunctionPointer(
                    pDeleteProps, typeof(DeletePropsDelegate));

                int hr = deleteProps(pMsg, pTagArray, IntPtr.Zero);
                if (hr != S_OK)
                {
                    logger?.Warn(LogComponent,
                        $"IMessage::DeleteProps for body props returned 0x{hr:X8}");
                    return false;
                }

                logger?.Debug(LogComponent, "Deleted PR_BODY, PR_BODY_HTML, PR_RTF_COMPRESSED");
                return true;
            }
            finally
            {
                if (pTagArray != IntPtr.Zero) Marshal.FreeCoTaskMem(pTagArray);
            }
        }

        /// <summary>
        /// Creates a hidden MAPI attachment on the IMessage containing the CMS
        /// envelope bytes. This is how Outlook natively stores S/MIME content:
        ///
        ///   PR_ATTACH_METHOD       = ATTACH_BY_VALUE (1)
        ///   PR_ATTACH_FILENAME     = "smime.p7m"
        ///   PR_ATTACH_LONG_FILENAME = "smime.p7m"
        ///   PR_ATTACH_MIME_TAG     = "application/pkcs7-mime; smime-type=enveloped-data"
        ///   PR_ATTACH_DATA_BIN     = [CMS envelope bytes]
        ///   PR_RENDERING_POSITION  = -1 (hidden — not in attachment well)
        /// </summary>
        private static bool CreateSmimeAttachment(
            IntPtr pMsg, byte[] cmsEnvelope, ParclLogger? logger)
        {
            IntPtr pAttach = IntPtr.Zero;

            try
            {
                // IMessage::CreateAttach(NULL, 0, &attachNum, &pAttach)
                IntPtr vtable = Marshal.ReadIntPtr(pMsg);
                IntPtr pCreateAttach = Marshal.ReadIntPtr(vtable,
                    VtableIndex_CreateAttach * IntPtr.Size);
                var createAttach = (CreateAttachDelegate)Marshal.GetDelegateForFunctionPointer(
                    pCreateAttach, typeof(CreateAttachDelegate));

                int hr = createAttach(pMsg, IntPtr.Zero, 0, out uint attachNum, out pAttach);
                if (hr != S_OK || pAttach == IntPtr.Zero)
                {
                    logger?.Error(LogComponent,
                        $"IMessage::CreateAttach failed: 0x{hr:X8}");
                    return false;
                }

                logger?.Debug(LogComponent, $"Created attachment #{attachNum}");

                // Set attachment properties via HrSetOneProp on the IAttach object

                // PR_ATTACH_METHOD = ATTACH_BY_VALUE
                if (!SetIntPropertyDirect(pAttach, PR_ATTACH_METHOD, ATTACH_BY_VALUE, logger))
                    return false;

                // PR_RENDERING_POSITION = -1 (hidden)
                if (!SetIntPropertyDirect(pAttach, PR_RENDERING_POSITION, -1, logger))
                    return false;

                // PR_ATTACH_FILENAME = "smime.p7m"
                if (!SetStringPropertyDirect(pAttach, PR_ATTACH_FILENAME_W, "smime.p7m", logger))
                    return false;

                // PR_ATTACH_LONG_FILENAME = "smime.p7m"
                if (!SetStringPropertyDirect(pAttach, PR_ATTACH_LONG_FILENAME_W, "smime.p7m", logger))
                    return false;

                // PR_ATTACH_MIME_TAG = "application/pkcs7-mime; smime-type=enveloped-data"
                if (!SetStringPropertyDirect(pAttach, PR_ATTACH_MIME_TAG_W,
                    "application/pkcs7-mime; smime-type=enveloped-data", logger))
                    return false;

                // PR_ATTACH_DATA_BIN = CMS envelope bytes
                if (!SetBinaryPropertyDirect(pAttach, PR_ATTACH_DATA_BIN, cmsEnvelope, logger))
                    return false;

                // SaveChanges on the attachment
                IntPtr attachVtable = Marshal.ReadIntPtr(pAttach);
                IntPtr pSaveChanges = Marshal.ReadIntPtr(attachVtable,
                    VtableIndex_SaveChanges * IntPtr.Size);
                var saveChanges = (SaveChangesDelegate)Marshal.GetDelegateForFunctionPointer(
                    pSaveChanges, typeof(SaveChangesDelegate));

                hr = saveChanges(pAttach, 0);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"IAttach::SaveChanges failed: 0x{hr:X8}");
                    return false;
                }

                logger?.Debug(LogComponent,
                    $"S/MIME attachment saved ({cmsEnvelope.Length} bytes)");
                return true;
            }
            finally
            {
                if (pAttach != IntPtr.Zero) Marshal.Release(pAttach);
            }
        }

        /// <summary>
        /// Sets a string property directly on a raw IMAPIProp pointer
        /// (used for IAttach objects where we already have the IntPtr).
        /// </summary>
        private static bool SetStringPropertyDirect(
            IntPtr pMapiProp, uint propTag, string value, ParclLogger? logger)
        {
            IntPtr pStr = Marshal.StringToCoTaskMemUni(value);
            try
            {
                var prop = new SPropValue
                {
                    ulPropTag = propTag,
                    dwAlignPad = 0
                };
                prop.Value.lpszW = pStr;

                int hr = HrSetOneProp(pMapiProp, ref prop);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"HrSetOneProp(0x{propTag:X8}) failed on attachment: 0x{hr:X8}");
                    return false;
                }

                return true;
            }
            finally
            {
                Marshal.FreeCoTaskMem(pStr);
            }
        }

        /// <summary>
        /// Sets an integer property directly on a raw IMAPIProp pointer.
        /// </summary>
        private static bool SetIntPropertyDirect(
            IntPtr pMapiProp, uint propTag, int value, ParclLogger? logger)
        {
            var prop = new SPropValue
            {
                ulPropTag = propTag,
                dwAlignPad = 0
            };
            prop.Value.l = value;

            int hr = HrSetOneProp(pMapiProp, ref prop);
            if (hr != S_OK)
            {
                logger?.Error(LogComponent,
                    $"HrSetOneProp(0x{propTag:X8}) failed on attachment: 0x{hr:X8}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Sets a binary property directly on a raw IMAPIProp pointer using
        /// the vtable SetProps method (needed for PT_BINARY on x64).
        /// </summary>
        private static bool SetBinaryPropertyDirect(
            IntPtr pMapiProp, uint propTag, byte[] data, ParclLogger? logger)
        {
            IntPtr pData = IntPtr.Zero;
            IntPtr pPropArray = IntPtr.Zero;

            try
            {
                pData = Marshal.AllocCoTaskMem(data.Length);
                Marshal.Copy(data, 0, pData, data.Length);

                int propSize = IntPtr.Size == 8 ? 24 : 16;
                pPropArray = Marshal.AllocCoTaskMem(propSize);

                for (int i = 0; i < propSize; i++)
                    Marshal.WriteByte(pPropArray, i, 0);

                Marshal.WriteInt32(pPropArray, 0, unchecked((int)propTag));

                if (IntPtr.Size == 8)
                {
                    Marshal.WriteInt32(pPropArray, 8, data.Length);
                    Marshal.WriteIntPtr(pPropArray, 16, pData);
                }
                else
                {
                    Marshal.WriteInt32(pPropArray, 8, data.Length);
                    Marshal.WriteIntPtr(pPropArray, 12, pData);
                }

                IntPtr vtable = Marshal.ReadIntPtr(pMapiProp);
                IntPtr pSetProps = Marshal.ReadIntPtr(vtable,
                    VtableIndex_SetProps * IntPtr.Size);
                var setProps = (SetPropsDelegate)Marshal.GetDelegateForFunctionPointer(
                    pSetProps, typeof(SetPropsDelegate));

                int hr = setProps(pMapiProp, 1, pPropArray, IntPtr.Zero);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"SetProps(0x{propTag:X8}) failed on attachment: 0x{hr:X8}");
                    return false;
                }

                return true;
            }
            finally
            {
                if (pData != IntPtr.Zero) Marshal.FreeCoTaskMem(pData);
                if (pPropArray != IntPtr.Zero) Marshal.FreeCoTaskMem(pPropArray);
            }
        }

        /// <summary>
        /// Calls IMessage::SaveChanges(0) to persist all property changes.
        /// </summary>
        private static bool CallSaveChanges(IntPtr pMsg, ParclLogger? logger)
        {
            try
            {
                IntPtr vtable = Marshal.ReadIntPtr(pMsg);
                IntPtr pSaveChanges = Marshal.ReadIntPtr(vtable,
                    VtableIndex_SaveChanges * IntPtr.Size);
                var saveChanges = (SaveChangesDelegate)Marshal.GetDelegateForFunctionPointer(
                    pSaveChanges, typeof(SaveChangesDelegate));

                int hr = saveChanges(pMsg, 0);
                if (hr != S_OK)
                {
                    logger?.Error(LogComponent,
                        $"IMessage::SaveChanges failed: 0x{hr:X8}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                logger?.Error(LogComponent, "Exception in SaveChanges", ex);
                return false;
            }
        }
    }
}
