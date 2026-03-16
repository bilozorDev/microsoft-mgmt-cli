export interface AuthMethod {
  Id: string;
  ODataType: string | null;
}

export const MFA_METHOD_NAMES: Record<string, string> = {
  microsoftAuthenticatorAuthenticationMethod: "Authenticator App",
  phoneAuthenticationMethod: "Phone",
  fido2AuthenticationMethod: "FIDO2 Security Key",
  emailAuthenticationMethod: "Email",
  softwareOathAuthenticationMethod: "Software Token",
  windowsHelloForBusinessAuthenticationMethod: "Windows Hello",
  temporaryAccessPassAuthenticationMethod: "Temporary Access Pass",
  platformCredentialAuthenticationMethod: "Platform Credential",
};

/** Per-type detail fetchers — each returns a human-readable detail string. */
export const MFA_DETAIL_CMDS: Record<string, { cmd: (uid: string, mid: string) => string; format: (raw: Record<string, unknown>) => string }> = {
  microsoftAuthenticatorAuthenticationMethod: {
    cmd: (uid, mid) =>
      `Get-MgUserAuthenticationMicrosoftAuthenticatorMethod -UserId '${uid}' -MicrosoftAuthenticatorAuthenticationMethodId '${mid}' | Select-Object DisplayName,DeviceTag,PhoneAppVersion,CreatedDateTime`,
    format: (r) => [r.DisplayName, r.DeviceTag, r.PhoneAppVersion ? `v${r.PhoneAppVersion}` : null].filter(Boolean).join(", "),
  },
  phoneAuthenticationMethod: {
    cmd: (uid, mid) =>
      `Get-MgUserAuthenticationPhoneMethod -UserId '${uid}' -PhoneAuthenticationMethodId '${mid}' | Select-Object PhoneNumber,PhoneType`,
    format: (r) => [r.PhoneNumber, r.PhoneType].filter(Boolean).join(" "),
  },
  fido2AuthenticationMethod: {
    cmd: (uid, mid) =>
      `Get-MgUserAuthenticationFido2Method -UserId '${uid}' -Fido2AuthenticationMethodId '${mid}' | Select-Object DisplayName,Model,CreatedDateTime`,
    format: (r) => [r.DisplayName, r.Model].filter(Boolean).join(", "),
  },
  emailAuthenticationMethod: {
    cmd: (uid, mid) =>
      `Get-MgUserAuthenticationEmailMethod -UserId '${uid}' -EmailAuthenticationMethodId '${mid}' | Select-Object EmailAddress`,
    format: (r) => String(r.EmailAddress ?? ""),
  },
};

export const MFA_REMOVE_CMDLETS: Record<string, { cmdlet: string; param: string }> = {
  microsoftAuthenticatorAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationMicrosoftAuthenticatorMethod",
    param: "-MicrosoftAuthenticatorAuthenticationMethodId",
  },
  phoneAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationPhoneMethod",
    param: "-PhoneAuthenticationMethodId",
  },
  fido2AuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationFido2Method",
    param: "-Fido2AuthenticationMethodId",
  },
  softwareOathAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationSoftwareOathMethod",
    param: "-SoftwareOathAuthenticationMethodId",
  },
  emailAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationEmailMethod",
    param: "-EmailAuthenticationMethodId",
  },
  windowsHelloForBusinessAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationWindowsHelloForBusinessMethod",
    param: "-WindowsHelloForBusinessAuthenticationMethodId",
  },
  temporaryAccessPassAuthenticationMethod: {
    cmdlet: "Remove-MgUserAuthenticationTemporaryAccessPassMethod",
    param: "-TemporaryAccessPassAuthenticationMethodId",
  },
};

export function friendlyMfaMethod(odataType: string): string | null {
  const lastSegment = odataType.split(".").pop() ?? "";
  if (lastSegment === "passwordAuthenticationMethod") return null;
  return MFA_METHOD_NAMES[lastSegment] ?? lastSegment;
}

export function mfaTypeKey(odataType: string): string {
  return odataType.split(".").pop() ?? "";
}
