using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace LyncLogger
{
	internal static class SecureCredentials
	{
		static readonly byte[] Entropy = System.Text.Encoding.Unicode.GetBytes("Some Random Salt 12389urnn232958y220r829f29n20389rh-5");

		public static string EncryptString(System.Security.SecureString input)
		{
			byte[] encryptedData = System.Security.Cryptography.ProtectedData.Protect(
				System.Text.Encoding.Unicode.GetBytes(ToInsecureString(input)),
				Entropy,
				System.Security.Cryptography.DataProtectionScope.CurrentUser);
			return Convert.ToBase64String(encryptedData);
		}

		public static SecureString DecryptString(string encryptedData)
		{
			try
			{
				byte[] decryptedData = System.Security.Cryptography.ProtectedData.Unprotect(
					Convert.FromBase64String(encryptedData),
					Entropy,
					System.Security.Cryptography.DataProtectionScope.CurrentUser);
				return ToSecureString(System.Text.Encoding.Unicode.GetString(decryptedData));
			}
			catch
			{
				return new SecureString();
			}
		}

		public static SecureString ToSecureString(string input)
		{
			SecureString secure = new SecureString();
			foreach (char c in input)
			{
				secure.AppendChar(c);
			}
			secure.MakeReadOnly();
			return secure;
		}

		public static string ToInsecureString(SecureString input)
		{
			string returnValue = string.Empty;
			IntPtr ptr = System.Runtime.InteropServices.Marshal.SecureStringToBSTR(input);
			try
			{
				returnValue = System.Runtime.InteropServices.Marshal.PtrToStringBSTR(ptr);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ZeroFreeBSTR(ptr);
			}
			return returnValue;
		}
	}
}
