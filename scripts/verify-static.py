#!/usr/bin/env python3
"""Portable static checks; no Office, VSTO, .NET, or MSI execution is attempted."""

import csv
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
INSTALLER = ROOT / "installer" / "Suite.Installer"
WIX_NS = "http://schemas.microsoft.com/wix/2006/wi"
MSBUILD_NS = "http://schemas.microsoft.com/developer/msbuild/2003"
GUID_RE = re.compile(r"^\{[0-9A-Fa-f]{8}(?:-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}\}$")
errors = []


def require(condition, message):
    if not condition:
        errors.append(message)


def parse_xml(path):
    try:
        return ET.parse(str(path)).getroot()
    except (ET.ParseError, OSError) as exc:
        errors.append("XML %s: %s" % (path.relative_to(ROOT), exc))
        return None


def check_powershell_structure(text):
    """Check comments, quoted strings, and paired delimiters without PowerShell."""
    pairs = {")": "(", "]": "[", "}": "{"}
    stack = []
    state = "normal"
    line = 1
    index = 0
    while index < len(text):
        char = text[index]
        following = text[index + 1] if index + 1 < len(text) else ""

        if state == "line_comment":
            if char == "\n":
                state = "normal"
        elif state == "block_comment":
            if char == "#" and following == ">":
                state = "normal"
                index += 2
                continue
        elif state == "single_quote":
            if char == "'" and following == "'":
                index += 2
                continue
            if char == "'":
                state = "normal"
        elif state == "double_quote":
            if char == "`":
                if following == "\n":
                    line += 1
                index += 2
                continue
            if char == '"':
                state = "normal"
        else:
            if char == "<" and following == "#":
                state = "block_comment"
                index += 2
                continue
            if char == "#":
                state = "line_comment"
            elif char == "'":
                state = "single_quote"
            elif char == '"':
                state = "double_quote"
            elif char in "([":
                stack.append((char, line))
            elif char == "{":
                stack.append((char, line))
            elif char in pairs:
                if not stack or stack[-1][0] != pairs[char]:
                    errors.append("PowerShell line %d: unmatched %s" % (line, char))
                    return
                stack.pop()

        if char == "\n":
            line += 1
        index += 1

    if state in ("block_comment", "single_quote", "double_quote"):
        errors.append("PowerShell: unterminated %s" % state)
    for delimiter, opened_at in stack:
        errors.append("PowerShell line %d: unclosed %s" % (opened_at, delimiter))


def check_csv():
    path = ROOT / "requirements" / "PublicFeatureMatrix.csv"
    required_columns = {
        "FeatureID", "CanonicalName", "Host", "SourceProducts", "SourceURLs",
        "Retrieved", "Status", "Dependency", "Phase", "VerificationCase", "Notes",
    }
    identifiers = set()
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            reader = csv.DictReader(stream)
            require(reader.fieldnames is not None, "CSV has no header")
            if reader.fieldnames is None:
                return
            require(required_columns.issubset(set(reader.fieldnames)), "CSV is missing required columns")
            row_count = 0
            for line_number, row in enumerate(reader, start=2):
                row_count += 1
                require(None not in row, "CSV line %d has extra columns" % line_number)
                require(all(value is not None for value in row.values()), "CSV line %d has missing columns" % line_number)
                feature_id = (row.get("FeatureID") or "").strip()
                require(bool(feature_id), "CSV line %d has no FeatureID" % line_number)
                require(feature_id not in identifiers, "CSV line %d duplicates %s" % (line_number, feature_id))
                identifiers.add(feature_id)
            require(row_count > 0, "CSV contains no feature rows")
    except (OSError, csv.Error) as exc:
        errors.append("CSV PublicFeatureMatrix.csv: %s" % exc)


def check_source_projects():
    project_paths = sorted((ROOT / "src").rglob("*.csproj")) + sorted((ROOT / "tests").rglob("*.csproj"))
    require(bool(project_paths), "No C# project files found")
    namespace = {"m": MSBUILD_NS}
    for path in project_paths:
        project = parse_xml(path)
        if project is None:
            continue
        framework = project.find(".//m:TargetFrameworkVersion", namespace)
        require(framework is not None and (framework.text or "").strip() == "v4.8",
                "%s must target .NET Framework 4.8" % path.relative_to(ROOT))

    required_tokens = {
        "src/Suite.Core/Data/ChineseIdCard.cs": ("IsValid", "TryGetBirthDate", "HasValidCheckCode"),
        "src/Suite.Core/Data/TabularData.cs": ("RemoveBlankRows", "DistinctRows", "Compare"),
        "src/Suite.Core/Batching/OutputPathPolicy.cs": ("EnsureDistinct", "GetAvailablePath"),
        "src/Suite.Core/Commands/CommandExecutor.cs": ("CanExecute", "OperationCanceledException"),
        "src/Suite.Infrastructure/DpapiSecretStore.cs": ("DataProtectionScope.CurrentUser", "ProtectedData.Protect"),
        "src/Suite.Infrastructure/HttpServiceConnector.cs": ("Uri.UriSchemeHttps", "ConfigureAwait(false)"),
    }
    for relative_path, tokens in required_tokens.items():
        path = ROOT / relative_path
        try:
            source = path.read_text(encoding="utf-8")
        except OSError as exc:
            errors.append("Source %s: %s" % (relative_path, exc))
            continue
        for token in tokens:
            require(token in source, "%s is missing key source marker %s" % (relative_path, token))


def check_word_command_ribbon_consistency():
    commands_path = ROOT / "src" / "Suite.Word" / "WordCommands.cs"
    ribbon_path = ROOT / "src" / "Suite.Word" / "WordRibbon.xml"
    try:
        source = commands_path.read_text(encoding="utf-8")
    except OSError as exc:
        errors.append("Word command consistency: %s" % exc)
        return
    ribbon = parse_xml(ribbon_path)
    if ribbon is None:
        return

    registered = set(re.findall(r'\bRegister(?:Context)?\(\s*"([A-Z0-9-]+)"', source))
    ribbon_tags = [node.get("tag", "") for node in ribbon.iter() if node.get("onAction") == "OnCommand"]
    require(len(ribbon_tags) == len(set(ribbon_tags)), "Word Ribbon contains duplicate command tags")
    require(all(tag in registered for tag in ribbon_tags),
            "Word Ribbon has unregistered command tags: %s" % sorted(set(ribbon_tags) - registered))

    extended = {
        "W-DOC-001-MERGE", "W-DOC-001-SPLIT-H1", "W-DOC-003-PDF",
        "W-TBL-001-FIT", "W-TBL-001-ROWS", "W-TBL-001-COLS", "W-LABEL-001-SEAT",
    }
    require(extended.issubset(registered), "Word extended commands missing: %s" % sorted(extended - registered))
    require(extended.issubset(set(ribbon_tags)), "Word extended Ribbon entries missing: %s" % sorted(extended - set(ribbon_tags)))


def check_excel_command_ribbon_consistency():
    commands_path = ROOT / "src" / "Suite.Excel" / "Commands" / "ExcelCommandSet.cs"
    ribbon_path = ROOT / "src" / "Suite.Excel" / "Ribbon" / "SuiteRibbon.xml"
    try:
        source = commands_path.read_text(encoding="utf-8")
    except OSError as exc:
        errors.append("Excel command consistency: %s" % exc)
        return
    ribbon = parse_xml(ribbon_path)
    if ribbon is None:
        return

    registered = set(re.findall(r'\bRegister\(\s*registry\s*,\s*"([A-Z0-9-]+)"', source))
    ribbon_tags = [node.get("tag", "") for node in ribbon.iter() if node.get("onAction") == "OnRibbonAction"]
    require(len(ribbon_tags) == len(set(ribbon_tags)), "Excel Ribbon contains duplicate command tags")
    require(all(tag in registered for tag in ribbon_tags),
            "Excel Ribbon has unregistered command tags: %s" % sorted(set(ribbon_tags) - registered))
    require({"E-MERGE-001-GROUPSUM", "E-EDU-001-SUMMARY"}.issubset(set(ribbon_tags)),
            "Excel extended Ribbon entries are missing")


def check_installer():
    wixproj_path = INSTALLER / "Suite.Installer.wixproj"
    product_path = INSTALLER / "Product.wxs"
    wixproj = parse_xml(wixproj_path)
    product_xml = parse_xml(product_path)
    try:
        product_text = product_path.read_text(encoding="utf-8")
    except OSError as exc:
        errors.append("Installer Product.wxs: %s" % exc)
        return

    if wixproj is not None:
        namespace = {"m": MSBUILD_NS}
        compile_item = wixproj.find(".//m:Compile[@Include='Product.wxs']", namespace)
        require(compile_item is not None, "WiX project does not compile Product.wxs")
        import_item = wixproj.find(".//m:Import", namespace)
        target_paths = wixproj.findall(".//m:WixTargetsPath", namespace)
        require(import_item is not None and import_item.get("Project") == "$(WixTargetsPath)" and
                any("v3.x" in (node.text or "") for node in target_paths),
                "WiX project must import WiX 3.x targets")
        constants = wixproj.find(".//m:DefineConstants", namespace)
        constant_text = "" if constants is None else (constants.text or "")
        for token in ("Configuration=$(Configuration)", "Platform=$(Platform)", "SuiteSourceDir=$(SuiteSourceDir)"):
            require(token in constant_text, "WiX project does not pass %s" % token)

    if product_xml is None:
        return
    namespace = {"w": WIX_NS}
    require(product_xml.tag == "{%s}Wix" % WIX_NS, "Product.wxs is not a WiX 3 document")
    product = product_xml.find("w:Product", namespace)
    require(product is not None, "Product.wxs has no Product")
    if product is None:
        return

    require(GUID_RE.match(product.get("UpgradeCode", "")) is not None,
            "Product UpgradeCode must be a fixed braced GUID")
    require(product.find("w:MajorUpgrade", namespace) is not None, "MajorUpgrade is missing")
    package = product.find("w:Package", namespace)
    require(package is not None, "Package is missing")
    if package is not None:
        require(package.get("InstallScope") == "perUser", "Package must install per user")
        require(package.get("InstallPrivileges") == "limited", "Package must use limited privileges")
        require(package.get("Platform") == "$(var.Platform)", "Package platform must be parameterized")

    for token in ('<?if $(var.Platform) = "x64" ?>', '<?if $(var.Platform) = "x86" ?>', 'Win64="$(var.Win64)"'):
        require(token in product_text, "Product.wxs is missing architecture handling: %s" % token)

    parents = {child: parent for parent in product_xml.iter() for child in parent}
    install_dir = product_xml.find(".//w:Directory[@Id='INSTALLDIR']", namespace)
    require(install_dir is not None, "INSTALLDIR is missing")
    if install_dir is not None:
        local_app_data = parents.get(install_dir)
        require(local_app_data is not None and local_app_data.get("Id") == "LocalAppDataFolder",
                "INSTALLDIR must be directly beneath LocalAppDataFolder")

    expected_files = {
        "Suite.Core.dll", "Suite.Infrastructure.dll", "Suite.Windowing.dll",
        "Suite.Word.dll", "Suite.Word.vsto", "Suite.Word.dll.manifest",
        "Suite.Excel.dll", "Suite.Excel.vsto", "Suite.Excel.dll.manifest",
        "Suite.PowerPoint.dll", "Suite.PowerPoint.vsto", "Suite.PowerPoint.dll.manifest",
    }
    components = product_xml.findall(".//w:Component", namespace)
    component_ids = set()
    component_guids = set()
    actual_files = set()
    for component in components:
        component_id = component.get("Id", "")
        guid = component.get("Guid", "")
        require(component_id and component_id not in component_ids, "Duplicate or empty Component Id: %s" % component_id)
        require(GUID_RE.match(guid) is not None, "Component %s needs a fixed valid GUID" % component_id)
        require(guid not in component_guids, "Component %s reuses a GUID" % component_id)
        component_ids.add(component_id)
        component_guids.add(guid)
        require(component.get("Win64") == "$(var.Win64)", "Component %s lacks parameterized Win64" % component_id)
        for file_node in component.findall("w:File", namespace):
            name = file_node.get("Name", "")
            source = file_node.get("Source", "")
            actual_files.add(name)
            require(file_node.get("KeyPath") == "yes", "File %s must be its component key path" % name)
            require(source.endswith("\\" + name), "File %s has an inconsistent source path" % name)
            lowered = name.lower()
            require(not lowered.startswith("microsoft.office.interop."),
                    "Office PIA must not be packaged: %s" % name)
            require(lowered not in {"office.dll", "vstoinstaller.exe", "microsoft.visualstudio.tools.office.runtime.dll"},
                    "Office/VSTO runtime must not be packaged: %s" % name)
    require(actual_files == expected_files,
            "Packaged file set differs; missing=%s extra=%s" %
            (sorted(expected_files - actual_files), sorted(actual_files - expected_files)))

    feature_refs = {node.get("Id") for node in product_xml.findall(".//w:Feature/w:ComponentRef", namespace)}
    require(feature_refs == component_ids, "Every component must be referenced by the product feature exactly by Id")

    registry_values = product_xml.findall(".//w:RegistryValue", namespace)
    registry_map = {(node.get("Root"), node.get("Key"), node.get("Name")): node for node in registry_values}
    for host in ("Word", "Excel", "PowerPoint"):
        key = r"Software\Microsoft\Office\%s\Addins\Suite.%s" % (host, host)
        for name in ("FriendlyName", "Description", "LoadBehavior", "Manifest"):
            require(("HKCU", key, name) in registry_map, "Missing HKCU %s %s registration" % (host, name))
        load_behavior = registry_map.get(("HKCU", key, "LoadBehavior"))
        if load_behavior is not None:
            require(load_behavior.get("Type") == "integer" and load_behavior.get("Value") == "3",
                    "%s LoadBehavior must be integer 3" % host)
        manifest = registry_map.get(("HKCU", key, "Manifest"))
        if manifest is not None:
            require(manifest.get("Value") == "file:///[INSTALLDIR]Suite.%s.vsto" % host,
                    "%s Manifest must target its installed .vsto file" % host)

    for value in registry_values:
        key = (value.get("Key") or "").lower()
        require("trustcenter" not in key and "\\security" not in key and "dontpromptforaddin" not in key,
                "Security/trust settings must not be changed: %s" % key)
        if "\\addins\\" in key:
            require(value.get("Root") == "HKCU", "Add-in registration must be per-user HKCU")
    require("vstolocal" not in product_text.lower(), "Do not suppress the normal VSTO/ClickOnce trust flow")
    require(not any("DownloadUrl" in node.attrib for node in product_xml.iter()),
            "Runtime downloads must not use unverified addresses")

    conditions = product_xml.findall(".//w:Condition", namespace)
    condition_text = " ".join("".join(node.itertext()) for node in conditions)
    require("NETFRAMEWORK48RELEASE" in condition_text and "#528040" in condition_text,
            "A blocking .NET Framework 4.8 launch condition is required")
    require("VSTORUNTIMEVERSION" in condition_text, "A blocking VSTO Runtime launch condition is required")
    for condition in conditions:
        message = condition.get("Message", "")
        require(bool(message.strip()), "Launch conditions must explain missing prerequisites")
        require("http://" not in message.lower() and "https://" not in message.lower(),
                "Launch conditions must not contain unverified download URLs")


def check_build_script():
    path = ROOT / "scripts" / "build-windows.ps1"
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        errors.append("PowerShell build-windows.ps1: %s" % exc)
        return
    check_powershell_structure(text)
    for token in (
        "vswhere.exe", "[17.0,18.0)", "MSBuild\\Current\\Bin\\MSBuild.exe",
        "/t:Restore", "/t:Build", "Suite.Core.Tests.exe", "Invoke-Native -FilePath $coreTestExe",
        "candle.exe", "light.exe", "signtool.exe", '"/pa"', "SUITE_SIGNING_PASSWORD", "ManifestCertificateThumbprint",
    ):
        require(token in text, "build-windows.ps1 is missing required step/tool: %s" % token)
    for forbidden in ("/noverify", "SkipSignatureVerification", "Set-ExecutionPolicy"):
        require(forbidden.lower() not in text.lower(), "build-windows.ps1 must not use %s" % forbidden)


def check_verification_script():
    path = ROOT / "scripts" / "verify-windows.ps1"
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        errors.append("PowerShell verify-windows.ps1: %s" % exc)
        return
    check_powershell_structure(text)
    for token in (
        "Get-RegistryValue", "SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full",
        "SOFTWARE\\Microsoft\\VSTO Runtime Setup\\v4R",
        "Microsoft.VisualStudio.OfficeTools.targets", "verify-static.py",
        "Suite.Core.Tests.exe", "-InstallMsi", "msiexec.exe", "/qn",
        "/norestart", "LoadBehavior", "Suite.Core.dll", "exit 1",
    ):
        require(token in text, "verify-windows.ps1 is missing required step/token: %s" % token)
    for forbidden in ("Set-ExecutionPolicy", "/noverify", "TrustCenter", "DisableAntiVirus"):
        require(forbidden.lower() not in text.lower(),
                "verify-windows.ps1 must not use %s" % forbidden)


def main():
    check_csv()
    check_source_projects()
    check_word_command_ribbon_consistency()
    check_excel_command_ribbon_consistency()
    check_installer()
    check_build_script()
    check_verification_script()
    if errors:
        for error in errors:
            print("FAIL: %s" % error, file=sys.stderr)
        print("Static verification failed with %d error(s)." % len(errors), file=sys.stderr)
        return 1
    print("Static verification passed: XML, CSV, key source markers, installer policy, and PowerShell structure.")
    print("No build, MSI installation, Office, VSTO Runtime, or Authenticode execution was performed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
