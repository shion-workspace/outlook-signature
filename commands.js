Office.onReady(() => {
    // Office ist bereit
});

// Diese Funktion wird automatisch beim Öffnen einer neuen E-Mail ausgelöst
async function onMessageComposeHandler(event) {
    try {
        // 1. SSO-Token vom angemeldeten Outlook-User holen
        const bootstrapToken = await Office.auth.getAccessToken({
            allowSignInPrompt: false,
            allowConsentPrompt: false,
            forMSGraphAccess: true
        });

        // 2. Benutzerdaten aus Microsoft Graph abrufen
        const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mail,businessPhones,mobilePhone", {
            headers: {
                "Authorization": `Bearer ${bootstrapToken}`
            }
        });

        let userData = {
            displayName: "",
            jobTitle: "",
            mail: "info@finanzkaufleute.de",
            phone: "(09522) 30436-0"
        };

        if (response.ok) {
            const profile = await response.json();
            userData.displayName = profile.displayName || "";
            userData.jobTitle = profile.jobTitle || "";
            userData.mail = profile.mail || "info@finanzkaufleute.de";
            if (profile.businessPhones && profile.businessPhones.length > 0) {
                userData.phone = profile.businessPhones[0];
            }
        }

        // 3. Signatur-HTML generieren und einfügen
        const signatureHtml = buildSignatureHtml(userData);
        setSignature(signatureHtml, event);

    } catch (error) {
        // Fallback, falls Graph/SSO mal blockiert ist: Allgemeine Firmensignatur einfügen
        console.error("Signatur-Fehler, nutze Fallback:", error);
        const fallbackHtml = buildSignatureHtml({
            displayName: "",
            jobTitle: "",
            mail: "info@finanzkaufleute.de",
            phone: "(09522) 30436-0"
        });
        setSignature(fallbackHtml, event);
    }
}

function setSignature(html, event) {
    Office.context.mailbox.item.body.setSignatureAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
            event.completed();
        }
    );
}

function buildSignatureHtml(user) {
    const logoUrl = "https://www.finanzkaufleute.de/wp-content/uploads/vfag_logo.png";
    const fbIconUrl = "https://www.finanzkaufleute.de/wp-content/uploads/2017/06/facebook.png";
    const xingIconUrl = "https://www.finanzkaufleute.de/wp-content/uploads/2017/06/xing.png";
    const ytIconUrl = "https://www.finanzkaufleute.de/wp-content/uploads/2017/06/youtube.png";
    const igIconUrl = "https://www.finanzkaufleute.de/wp-content/uploads/2026/11/instagram-e1795691716909.jpg"
    const bannerUrl = "https://shion-workspace.github.io/outlook-signature/partnertag_email_banner.jpg"

    // Optionaler persönlicher Block (erscheint nur, wenn im M365-Profil ein Name hinterlegt ist)
    const personalBlock = user.displayName ? `
        <p style="margin: 0 0 12px 0; font-family: Arial, Helvetica, sans-serif; font-size: 13px; color: #222222;">
            <strong>${user.displayName}</strong>${user.jobTitle ? `<br><span style="color: #555555;">${user.jobTitle}</span>` : ""}
        </p>
    ` : "";

    return `
    <div style="font-family: Arial, Helvetica, sans-serif; font-size: 12px; color: #222222; line-height: 1.45;">
        <p style="margin: 0 0 12px 0; color: #555555;">-------------------------------------------</p>
        
        ${personalBlock}

        <p style="margin: 0 0 8px 0; font-size: 13px;">
            <strong>Vereinigte Finanzkaufleute<br>in Deutschland AG</strong><br>
            Adresse: Raiffeisenstr. 5 | 97514 Oberaurach<br>
            Tel: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${user.phone}<br>
            E-Mail: &nbsp;&nbsp;&nbsp;<a href="mailto:${user.mail}" style="color: #004b87; text-decoration: underline;">${user.mail}</a><br>
            Internet: <a href="https://www.finanzkaufleute.de" style="color: #004b87; text-decoration: underline;">www.finanzkaufleute.de</a>
        </p>

        <hr style="border: 0; border-top: 1px solid #cccccc; width: 280px; margin: 12px 0; text-align: left;" />

        <p style="margin: 10px 0;">
            <a href="https://www.finanzkaufleute.de" target="https://www.finanzkaufleute.de" style="text-decoration: none;">
                <img src="${logoUrl}" alt="Vereinigte Finanzkaufleute in Deutschland AG" style="max-width: 260px; height: auto; border: 0; display: block;" />
            </a>
        </p>

        <!-- Social Media Icons -->
        <p style="margin: 10px 0 16px 0;">
            <a href="https://www.facebook.com/vereinigtefinanzkaufleute" target="_blank" style="text-decoration: none; margin-right: 6px;">
                <img src="${fbIconUrl}" alt="Facebook" width="20" height="20" style="border: 0; vertical-align: middle;" />
            </a>
            <a href="https://www.xing.com/pages/vereinigtefinanzkaufleuteindeutschlandag" target="_blank" style="text-decoration: none; margin-right: 6px;">
                <img src="${xingIconUrl}" alt="Xing" width="20" height="20" style="border: 0; vertical-align: middle;" />
            </a>
            <a href="https://www.youtube.com/user/Finanzkaufleute" target="_blank" style="text-decoration: none;">
                <img src="${ytIconUrl}" alt="YouTube" width="20" height="20" style="border: 0; vertical-align: middle;" />
            </a>
            <a href="https://www.instagram.com/vereinigte_finanzkaufleute" target="_blank" style="text-decoration: none; margin-right: 6px;">
                <img src="${igIconUrl}" alt="Intagram" width="20" height="20" style="border: 0; vertical-align: middle;" />
            </a>
        </p>

        <p style="margin: 12px 0 16px 0;">
            <a href="https://www.finanzkaufleute.de/vfag-partnertag/partnertag-informationen/" target="_blank" style="text-decoration: none">
                <img src="${bannerUrl}" alt="VFAG Partnertag 2027" style="max-width: 550px; width: 100%; height: auto; border: 0; display: block;"/>
            </a>
        </p>

        <!-- Pflichtangaben & Rechtstexte -->
        <div style="font-size: 10px; color: #555555; line-height: 1.4;">
            <p style="margin: 0 0 8px 0;">Vorstand: Jasmin Schmittwolf</p>
            <p style="margin: 0 0 8px 0;">Vorsitzende des Aufsichtsrats: Sabrina Elbin | Stellvertretender Vorsitzender des Aufsichtsrats: Thorsten Schmittwolf</p>
            <p style="margin: 0 0 8px 0;">Inhaltlich Verantwortlicher gemäß § 55 Abs. 2 RStV: Jasmin Schmittwolf | Raiffeisenstr. 5 | 97514 Oberaurach</p>
            <p style="margin: 0 0 8px 0;">Registergericht: Amtsgericht Bamberg | Handelsregister: B 36 21 | Umsatzsteueridentifikationsnummer: DE233614284</p>
            
            <p style="margin: 0 0 8px 0;">
                Erlaubniserteilung nach § 34 c GewO (Privatdarlehensvermittler): IHK München und Oberbayern | Max-Joseph-Str. 2 | 80333 München | Internet: <a href="https://www.ihk-muenchen.de" style="color: #555555;">www.ihk-muenchen.de</a> |<br>
                Erlaubniserteilung nach § 34 d Abs. 1 GewO (Versicherungsmakler): IHK München und Oberbayern | Max-Joseph-Str. 2 | 80333 München | Internet: <a href="https://www.ihk-muenchen.de" style="color: #555555;">www.ihk-muenchen.de</a> | Registriernummer D-Q94J-XZZ20-22<br>
                Erlaubniserteilung nach § 34 f Abs. 1 S. 1 Nr. 1-3 GewO (Finanzanlagenvermittler): IHK München und Oberbayern | Max-Joseph-Str. 2 | 80333 München | Internet: <a href="https://www.ihk-muenchen.de" style="color: #555555;">www.ihk-muenchen.de</a> | Registriernummer D-F-155-2FJV-67<br>
                Erlaubniserteilung nach § 34 i Abs. 1 S. 1 GewO (Immobiliardarlehensvermittler): IHK München und Oberbayern | Max-Joseph-Str. 2 | 80333 München | Internet: <a href="https://www.ihk-muenchen.de" style="color: #555555;">www.ihk-muenchen.de</a> | Registriernummer D-W-155-DPYA-74
            </p>

            <p style="margin: 0 0 8px 0;">
                Eingetragen im Versicherungsvermittlerregister: Deutscher Industrie- und Handelskammertag (DIHK) e. V. | Breite Straße 29 | 10178 Berlin<br>
                Telefon: +49 (0)180 600 58 50 (Festnetzpreis 0,20 €/Anruf; Mobilfunkpreise maximal 0,60 €/Anruf) | Internet: <a href="https://www.vermittlerregister.info" style="color: #555555;">www.vermittlerregister.info</a>
            </p>

            <p style="margin: 0 0 8px 0;">Zuständige Aufsichtsbehörde: Bundesanstalt für Finanzdienstleistungsaufsicht (BaFin) | Graurheindorfer Straße 108 | 53117 Bonn</p>

            <p style="margin: 0 0 8px 0;">
                Anschriften von Schlichtungsstellen:<br>
                Versicherungsombudsmann e.V.: Postfach 08 06 32 | 10006 Berlin | Internet: <a href="https://www.versicherungsombudsmann.de" style="color: #555555;">www.versicherungsombudsmann.de</a><br>
                Ombudsmann für die private Kranken- und Pflegeversicherung: Postfach 06 02 22, 10052 Berlin | Internet: <a href="https://www.pkv-ombudsmann.de" style="color: #555555;">www.pkv-ombudsmann.de</a><br>
                Ombudsmann für Investmentfonds des BVI | Unter den Linden 42 | 10117 Berlin | Internet: <a href="https://www.ombudsstelle-investmentfonds.de" style="color: #555555;">www.ombudsstelle-investmentfonds.de</a>
            </p>

            <p style="margin: 8px 0 0 0;">
                Hinweis: Der Inhalt dieser E-Mail ist ausschließlich für den bezeichneten Empfänger bestimmt und kann vertrauliche Informationen enthalten. Wenn Sie nicht der Empfänger dieser E-Mail oder dessen Vertreter sind, dann senden Sie bitte die E-Mail an den Absender zurück (Antwort-Funktion). Entfernen Sie bitte danach die Nachricht aus Ihrem System.
            </p>
        </div>
    </div>
    `;
}

// Verknüpfung für das Outlook-Event registrieren
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
