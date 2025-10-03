using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Forms = System.Windows.Forms;

namespace JsonToExcel
{
    public partial class MainWindow : Window
    {
        private readonly string _configPath;
        private bool _apiSync;
        private bool _apiVisible;

        public MainWindow()
        {
            InitializeComponent();
            _configPath = System.IO.Path.Combine(AppContext.BaseDirectory, "config.json");
            this.Loaded += Window_Loaded;
            // Afficher la clé API en texte normal par défaut (pas en mode mot de passe)
            _apiVisible = true;
            ApplyApiVisibility();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            LoadConfig();
        }

        private void SaveApiButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var cfg = LoadConfigInternal() ?? new AppConfig();
                cfg.EndpointUrl = EndpointTextBox.Text?.Trim() ?? string.Empty;
                cfg.ApiKey = GetApiKey();
                SaveConfigInternal(cfg);
                AppendLog("Paramètres API sauvegardés.");
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur lors de la sauvegarde API: {ex.Message}");
            }
        }

        private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using var dialog = new Forms.FolderBrowserDialog
                {
                    Description = "Choisissez le dossier de sortie",
                    ShowNewFolderButton = true
                };
                var result = dialog.ShowDialog();
                if (result == Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    OutputFolderTextBox.Text = dialog.SelectedPath;
                    var cfg = LoadConfigInternal() ?? new AppConfig();
                    cfg.OutputFolder = dialog.SelectedPath;
                    cfg.DeleteOld = DeleteOldCheckBox.IsChecked == true;
                    SaveConfigInternal(cfg);
                    AppendLog("Dossier de sortie enregistré.");
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur lors du choix du dossier: {ex.Message}");
            }
        }

        private void DeleteOldCheckBox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var cfg = LoadConfigInternal() ?? new AppConfig();
                cfg.OutputFolder = OutputFolderTextBox.Text?.Trim() ?? string.Empty;
                cfg.DeleteOld = DeleteOldCheckBox.IsChecked == true;
                SaveConfigInternal(cfg);
                AppendLog("Préférence 'Supprimer anciens' sauvegardée.");
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur de sauvegarde: {ex.Message}");
            }
        }

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleUi(false);
            try
            {
                var cfg = LoadConfigInternal() ?? new AppConfig();

                string endpoint = (EndpointTextBox.Text ?? string.Empty).Trim();
                string apiKey = GetApiKey();
                string outDir = (OutputFolderTextBox.Text ?? string.Empty).Trim();
                bool deleteOld = DeleteOldCheckBox.IsChecked == true;

                if (string.IsNullOrWhiteSpace(endpoint))
                {
                    AppendLog("Endpoint manquant.");
                    return;
                }
                if (string.IsNullOrWhiteSpace(outDir))
                {
                    AppendLog("Dossier de sortie manquant.");
                    return;
                }

                Directory.CreateDirectory(outDir);

                if (deleteOld)
                {
                    AppendLog("Suppression des anciens fichiers...");
                    TryDeleteDirectoryContents(outDir);
                }

                AppendLog("Connexion à l’API...");
                using var http = new HttpClient();
                if (!string.IsNullOrWhiteSpace(apiKey))
                {
                    if (http.DefaultRequestHeaders.Contains("X-AUTH-TOKEN"))
                        http.DefaultRequestHeaders.Remove("X-AUTH-TOKEN");
                    http.DefaultRequestHeaders.Add("X-AUTH-TOKEN", apiKey);
                }

                var response = await http.GetAsync(endpoint);
                response.EnsureSuccessStatusCode();
                AppendLog("Connexion à l’API OK");

                string json = await response.Content.ReadAsStringAsync();
                var root = ParseJson(json);
                if (root == null)
                {
                    AppendLog("JSON invalide ou vide.");
                    return;
                }

                var articles = root["articles"] as JArray;
                if (articles != null)
                {
                    AppendLog($"Traitement des articles ({articles.Count})...");
                    await GenerateByDivisionAndFamille(outDir, articles);
                }
                else
                {
                    // Fallback héritée: gestion par 'categories'
                    var categoriesToken = root["categories"];
                    if (categoriesToken == null)
                    {
                        AppendLog("Aucune clé 'categories' trouvée, génération d’un Excel unique.");
                        string defaultName = "categorie";
                        await GenerateExcelForCategory(outDir, defaultName, root);
                    }
                    else if (categoriesToken is JArray arr)
                    {
                        int count = 0;
                        for (int i = 0; i < arr.Count; i++)
                        {
                            var cat = arr[i];
                            string catName = ResolveCategoryName(cat, i);
                            await GenerateExcelForCategory(outDir, catName, cat);
                            count++;
                        }
                        AppendLog($"{count} fichiers Excel générés dans {outDir}");
                    }
                    else
                    {
                        string catName = ResolveCategoryName(categoriesToken, 0);
                        await GenerateExcelForCategory(outDir, catName, categoriesToken);
                        AppendLog($"1 fichier Excel généré dans {outDir}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur: {ex.Message}");
            }
            finally
            {
                ToggleUi(true);
            }
        }

        private async Task GenerateByDivisionAndFamille(string outDir, JArray articles)
        {
            // Groupe par division/famille
            var groups = articles
                .OfType<JObject>()
                .GroupBy(a => new
                {
                    Division = (a["division"]?.Type == JTokenType.Null ? null : (string?)a["division"]) ?? "SansDivision",
                    Famille = (a["famille"]?.Type == JTokenType.Null ? null : (string?)a["famille"]) ?? "SansFamille"
                });

            int totalExcel = 0;
            var allRowsGlobal = new List<Dictionary<string, string>>();
            var allCaracColsGlobal = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var nonEmptyCaracColsGlobal = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var g in groups)
            {
                var divName = string.IsNullOrWhiteSpace(g.Key.Division) ? "SansDivision" : g.Key.Division.Trim();
                var famName = string.IsNullOrWhiteSpace(g.Key.Famille) ? "SansFamille" : g.Key.Famille.Trim();

                string divisionFolder = Path.Combine(outDir, SanitizeName(divName));
                string familleFolder = Path.Combine(divisionFolder, SanitizeName(famName));
                Directory.CreateDirectory(familleFolder);

                AppendLog($"Division: '{divName}' / Famille: '{famName}' → {g.Count()} articles");

                var rows = new List<Dictionary<string, string>>();
                var allCaracColsGroup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var nonEmptyCaracColsGroup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var article in g)
                {
                    var row = BuildTopLevelRow(article);
                    AddCaracteristiquesColumns(article, row, allCaracColsGroup, nonEmptyCaracColsGroup);
                    rows.Add(row);

                    // Alimente l'index global
                    allRowsGlobal.Add(row.ToDictionary(k => k.Key, v => v.Value));
                    var carac = article["caracteristiques"] as JArray;
                    if (carac != null)
                    {
                        foreach (var item in carac.OfType<JObject>())
                        {
                            var text = (string?)item["text"];
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                allCaracColsGlobal.Add(text);
                                var valueToken = item["value"];
                                if (valueToken != null && valueToken.Type != JTokenType.Null)
                                {
                                    var valueStr = valueToken.ToString(Formatting.None);
                                    if (!string.IsNullOrWhiteSpace(valueStr))
                                        nonEmptyCaracColsGlobal.Add(text);
                                }
                            }
                        }
                    }
                }

                string fileName = ToCamelCase(famName) + ".xlsx";
                string xlsxPath = Path.Combine(familleFolder, fileName);
                await WriteExcel(xlsxPath, famName, rows, allCaracColsGroup, nonEmptyCaracColsGroup);
                AppendLog($"Excel généré: {xlsxPath}");
                totalExcel++;
            }

            AppendLog($"{totalExcel} fichiers Excel générés dans {outDir}");

            // Génère l'index global nettoyé à la racine
            string indexPath = Path.Combine(outDir, "index.xlsx");
            await WriteExcel(indexPath, "index", allRowsGlobal, allCaracColsGlobal, nonEmptyCaracColsGlobal);
            AppendLog($"Index global généré: {indexPath}");
        }

        private static Dictionary<string, string> BuildTopLevelRow(JObject article)
        {
            var dict = new Dictionary<string, string>();
            foreach (var prop in article.Properties())
            {
                var name = prop.Name;
                if (string.Equals(name, "caracteristiques", StringComparison.OrdinalIgnoreCase)) continue;
                if (string.Equals(name, "documents", StringComparison.OrdinalIgnoreCase)) continue; // ignoré
                if (string.Equals(name, "dispo", StringComparison.OrdinalIgnoreCase)) continue; // ignoré
                if (string.Equals(name, "calibre", StringComparison.OrdinalIgnoreCase)) continue; // ignoré
                if (string.Equals(name, "puissance", StringComparison.OrdinalIgnoreCase)) continue; // ignoré

                var val = prop.Value;
                switch (val.Type)
                {
                    case JTokenType.Array:
                        if (val is JArray arr)
                        {
                            // Joindre les tableaux en chaîne lisible avec virgule (sans guillemets superflus)
                            var joined = string.Join(", ", arr.Select(x =>
                            {
                                if (x.Type == JTokenType.Null) return string.Empty;
                                if (x is JValue v && v.Type == JTokenType.String) return NormalizeStringValue(v.Value<string>());
                                return x.ToString(Formatting.None);
                            }));
                            dict[name] = joined;
                        }
                        break;
                    case JTokenType.Object:
                        // Objet top-level: sérialiser compact
                        dict[name] = val.ToString(Formatting.None);
                        break;
                    case JTokenType.Null:
                        dict[name] = string.Empty;
                        break;
                    default:
                        if (val is JValue sv && sv.Type == JTokenType.String)
                        {
                            var s = NormalizeStringValue(sv.Value<string>());
                            if (string.Equals(name, "categorie", StringComparison.OrdinalIgnoreCase))
                                s = NormalizeCategorie(s);
                            dict[name] = s;
                        }
                        else
                        {
                            dict[name] = val.ToString(Formatting.None);
                        }
                        break;
                }
            }
            return dict;
        }

        private static void AddCaracteristiquesColumns(JObject article, Dictionary<string, string> dict, ISet<string>? allCaracColumns = null, ISet<string>? nonEmptyCaracColumns = null)
        {
            var carac = article["caracteristiques"] as JArray;
            if (carac == null || carac.Count == 0) return;
            foreach (var item in carac.OfType<JObject>())
            {
                var text = (string?)item["text"];
                if (string.IsNullOrWhiteSpace(text)) continue;
                if (IsMediaLabel(text)) continue; // ignorer "Médias"
                allCaracColumns?.Add(text);
                var valueToken = item["value"];
                if (valueToken == null || valueToken.Type == JTokenType.Null) continue;
                string value;
                if (valueToken is JValue jv && jv.Type == JTokenType.String)
                    value = NormalizeStringValue(jv.Value<string>());
                else
                    value = valueToken.ToString(Formatting.None);
                if (string.IsNullOrWhiteSpace(value)) continue;
                dict[text] = value;
                nonEmptyCaracColumns?.Add(text);
            }
        }

        private static Task WriteExcel(string path, string sheetName, List<Dictionary<string, string>> rows, ISet<string>? allCaracColumns = null, ISet<string>? nonEmptyCaracColumns = null)
        {
            var headers = rows.SelectMany(d => d.Keys).Distinct().ToList();
            if (allCaracColumns != null && nonEmptyCaracColumns != null)
            {
                headers = headers
                    .Where(h => !allCaracColumns.Contains(h) || nonEmptyCaracColumns.Contains(h))
                    .ToList();
            }

            var orderedHeaders = OrderHeaders(headers);
            var sortedRows = SortRowsForOutput(rows);

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add(SanitizeWorksheetName(string.IsNullOrWhiteSpace(sheetName) ? "Feuille1" : sheetName));

                // Header row
                var headerCaptions = orderedHeaders.Select(h => (h ?? string.Empty).ToUpperInvariant()).ToList();
                for (int c = 0; c < headerCaptions.Count; c++)
                {
                    ws.Cell(1, c + 1).Value = headerCaptions[c];
                    ws.Cell(1, c + 1).Style.Font.Bold = true;
                }

                // Data rows start at row 2
                for (int r = 0; r < sortedRows.Count; r++)
                {
                    var row = sortedRows[r];
                    for (int c = 0; c < orderedHeaders.Count; c++)
                    {
                        var key = orderedHeaders[c];
                        row.TryGetValue(key, out var val);
                        var headerNorm = NormalizeHeaderName(key);
                        if (headerNorm == "PUHT" || headerNorm == "PPC")
                        {
                            if (TryParseDecimalFlexible(val, out var dec))
                            {
                                ws.Cell(r + 2, c + 1).Value = dec;
                                ws.Cell(r + 2, c + 1).Style.NumberFormat.Format = "€ #,##0.00";
                            }
                            else
                            {
                                ws.Cell(r + 2, c + 1).Value = NormalizeStringValue(val);
                            }
                        }
                        else
                        {
                            ws.Cell(r + 2, c + 1).Value = NormalizeStringValue(val);
                        }
                    }
                }

                // Convertir en tableau structuré (table) couvrant toute la zone
                var rng = ws.Range(1, 1, Math.Max(1, sortedRows.Count) + 1, Math.Max(1, orderedHeaders.Count));
                var table = rng.CreateTable();
                table.Theme = XLTableTheme.None; // pas de thème automatique
                // Style de l'en-tête: fond rouge, texte blanc, gras, MAJUSCULES déjà appliquées
                var headerRow = table.HeadersRow();
                headerRow.Style.Fill.BackgroundColor = XLColor.Red;
                headerRow.Style.Font.FontColor = XLColor.White;
                headerRow.Style.Font.Bold = true;

                ws.Columns().AdjustToContents();

                using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    wb.SaveAs(fs);
                    fs.Flush(true);
                }
            }

            return Task.CompletedTask;
        }

        private static List<Dictionary<string, string>> SortRowsForOutput(IEnumerable<Dictionary<string, string>> rows)
        {
            string Norm(string? s) => NormalizeSortValue(s);
            return rows
                .OrderBy(r => Norm(GetValueCI(r, "division")))
                .ThenBy(r => Norm(GetValueCI(r, "famille")))
                .ThenBy(r => Norm(GetValueCI(r, "marque")))
                .ThenBy(r => Norm(GetValueCI(r, "nom")))
                .ToList();
        }

        private static string GetValueCI(Dictionary<string, string> row, string key)
        {
            foreach (var kvp in row)
            {
                if (string.Equals(kvp.Key, key, StringComparison.OrdinalIgnoreCase))
                    return kvp.Value ?? string.Empty;
            }
            return string.Empty;
        }

        private static string NormalizeSortValue(string? s)
        {
            var t = NormalizeStringValue(s);
            t = RemoveDiacritics(t);
            return t.ToUpperInvariant();
        }

        private static List<string> OrderHeaders(List<string> headers)
        {
            var preferred = new List<string>
            {
                "UNIVERS","DIVISION","FAMILLE","MARQUE","REFERENCE_FOURNISSEUR",
                "STATUS","STATUT","NOM","DESIGNATION","REFERENCE","CATEGORIE",
                "PUHT","PPC","DESCRIPTIF","RGA"
            };
            var used = new HashSet<string>();
            var result = new List<string>();

            // Helper to find and add header by normalized name
            foreach (var want in preferred)
            {
                var idx = headers.FindIndex(h => NormalizeHeaderName(h) == want);
                if (idx >= 0)
                {
                    var key = headers[idx];
                    if (used.Add(key)) result.Add(key);
                }
            }
            // Keep remaining headers in current order
            foreach (var h in headers)
            {
                if (used.Add(h)) result.Add(h);
            }
            return result;
        }

        private static string NormalizeHeaderName(string? name)
        {
            name ??= string.Empty;
            return RemoveDiacritics(name).ToUpperInvariant();
        }

        private static bool TryParseDecimalFlexible(string? input, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(input)) return false;
            var s = input.Trim();
            // remove surrounding quotes if any
            if (s.Length >= 2 && s[0] == '"' && s[^1] == '"') s = s.Substring(1, s.Length - 2);
            s = s.Replace(" \u00A0", " ");
            // Try invariant with dot
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
            // Replace comma by dot and try again
            var s2 = s.Replace(',', '.');
            if (decimal.TryParse(s2, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
            // Try fr-FR
            if (decimal.TryParse(s, NumberStyles.Any, new CultureInfo("fr-FR"), out value)) return true;
            return false;
        }

        private static string NormalizeStringValue(string? s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var t = s.Trim();
            // retire une paire de guillemets superflus si elle entoure tout le texte
            if (t.Length >= 2 && t.StartsWith("\"") && t.EndsWith("\""))
            {
                t = t.Substring(1, t.Length - 2);
            }
            return t;
        }

        private static string NormalizeCategorie(string? s)
        {
            var t = NormalizeStringValue(s);
            if (string.IsNullOrWhiteSpace(t)) return t;
            t = t.Trim();
            if (t.Length == 1) return t.ToUpperInvariant();
            return char.ToUpperInvariant(t[0]) + (t.Length > 1 ? t.Substring(1).ToLowerInvariant() : string.Empty);
        }

        private static bool IsMediaLabel(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var norm = RemoveDiacritics(text).ToLowerInvariant();
            return norm == "medias"; // match "Médias"
        }

        private static string RemoveDiacritics(string text)
        {
            var normalized = text.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();
            foreach (var ch in normalized)
            {
                var uc = System.Globalization.CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != System.Globalization.UnicodeCategory.NonSpacingMark)
                    sb.Append(ch);
            }
            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        private static string SanitizeWorksheetName(string name)
        {
            var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
            foreach (var ch in invalid)
            {
                name = name.Replace(ch, '_');
            }
            if (name.Length > 31) name = name.Substring(0, 31);
            if (string.IsNullOrWhiteSpace(name)) name = "Feuille1";
            return name;
        }

        private static string ToCamelCase(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var parts = new List<string>();
            var sb = new StringBuilder();
            foreach (var ch in input)
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(ch);
                else
                {
                    if (sb.Length > 0) { parts.Add(sb.ToString()); sb.Clear(); }
                }
            }
            if (sb.Length > 0) parts.Add(sb.ToString());

            if (parts.Count == 0) return string.Empty;
            var first = parts[0].ToLowerInvariant();
            var tail = parts.Skip(1).Select(p => p.Length == 0 ? string.Empty : char.ToUpperInvariant(p[0]) + (p.Length > 1 ? p.Substring(1).ToLowerInvariant() : string.Empty));
            return first + string.Concat(tail);
        }

        private void ToggleUi(bool enabled)
        {
            EndpointTextBox.IsEnabled = enabled;
            ApiKeyBox.IsEnabled = enabled;
            ApiKeyPlainTextBox.IsEnabled = enabled;
            ApiRevealToggle.IsEnabled = enabled;
            SaveApiButton.IsEnabled = enabled;
            OutputFolderTextBox.IsEnabled = enabled;
            BrowseFolderButton.IsEnabled = enabled;
            DeleteOldCheckBox.IsEnabled = enabled;
            GenerateButton.IsEnabled = enabled;
        }

        private void AppendLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogsTextBox.AppendText($"- {message}{Environment.NewLine}");
                LogsTextBox.ScrollToEnd();
            });
        }

        private void LoadConfig()
        {
            try
            {
                var cfg = LoadConfigInternal();
                if (cfg != null)
                {
                    if (!string.IsNullOrWhiteSpace(cfg.EndpointUrl)) EndpointTextBox.Text = cfg.EndpointUrl;
                    if (!string.IsNullOrWhiteSpace(cfg.ApiKey)) { ApiKeyBox.Password = cfg.ApiKey; ApiKeyPlainTextBox.Text = cfg.ApiKey; }
                    if (!string.IsNullOrWhiteSpace(cfg.OutputFolder)) OutputFolderTextBox.Text = cfg.OutputFolder;
                    DeleteOldCheckBox.IsChecked = cfg.DeleteOld;
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur lors du chargement de la configuration: {ex.Message}");
            }
        }

        private string GetApiKey()
        {
            return _apiVisible ? (ApiKeyPlainTextBox.Text ?? string.Empty) : (ApiKeyBox.Password ?? string.Empty);
        }

        private void ApiRevealToggle_Click(object sender, RoutedEventArgs e)
        {
            _apiVisible = !_apiVisible;
            ApplyApiVisibility();
        }

        private void ApplyApiVisibility()
        {
            if (_apiVisible)
            {
                // rendre visible le texte
                if (!_apiSync)
                {
                    _apiSync = true;
                    ApiKeyPlainTextBox.Text = ApiKeyBox.Password;
                    _apiSync = false;
                }
                ApiKeyPlainTextBox.Visibility = Visibility.Visible;
                ApiKeyBox.Visibility = Visibility.Collapsed;
                ApiRevealToggle.Content = "🙈"; // oeil barré = dots si recliqué
            }
            else
            {
                // masquer (dots)
                if (!_apiSync)
                {
                    _apiSync = true;
                    ApiKeyBox.Password = ApiKeyPlainTextBox.Text;
                    _apiSync = false;
                }
                ApiKeyPlainTextBox.Visibility = Visibility.Collapsed;
                ApiKeyBox.Visibility = Visibility.Visible;
                ApiRevealToggle.Content = "👁"; // oeil = clef lisible si cliqué
            }
        }

        private void ApiKeyBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (_apiSync) return;
            _apiSync = true;
            ApiKeyPlainTextBox.Text = ApiKeyBox.Password;
            _apiSync = false;
        }

        private void ApiKeyPlainTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (_apiSync) return;
            _apiSync = true;
            ApiKeyBox.Password = ApiKeyPlainTextBox.Text;
            _apiSync = false;
        }

        private AppConfig? LoadConfigInternal()
        {
            if (!File.Exists(_configPath)) return null;
            var json = File.ReadAllText(_configPath, Encoding.UTF8);
            return JsonConvert.DeserializeObject<AppConfig>(json);
        }

        private void SaveConfigInternal(AppConfig cfg)
        {
            var json = JsonConvert.SerializeObject(cfg, Formatting.Indented);
            File.WriteAllText(_configPath, json, Encoding.UTF8);
        }

        private static JObject? ParseJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return null;
            try
            {
                var token = JToken.Parse(json);
                if (token is JObject o) return o;
                // Si la racine est un tableau, l’envelopper sous une clé générique
                if (token is JArray a)
                {
                    return new JObject { ["categories"] = a };
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        private static string ResolveCategoryName(JToken cat, int index)
        {
            string? name = null;
            if (cat.Type == JTokenType.Object)
            {
                var o = (JObject)cat;
                name = (string?)o["name"] ?? (string?)o["category"] ?? (string?)o["title"] ?? (string?)o["id"];
            }
            else if (cat.Type == JTokenType.String)
            {
                name = (string?)cat;
            }
            if (string.IsNullOrWhiteSpace(name)) name = $"categorie_{index + 1}";
            return SanitizeName(name);
        }

        private static string SanitizeName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            return name.Trim();
        }

        private Task GenerateExcelForCategory(string outputDir, string categoryName, JToken data)
        {
            string categoryFolder = Path.Combine(outputDir, SanitizeName(categoryName));
            Directory.CreateDirectory(categoryFolder);
            string xlsxPath = Path.Combine(categoryFolder, $"{SanitizeName(categoryName)}.xlsx");

            List<Dictionary<string, string>> rows;
            if (data is JObject obj && obj["items"] is JArray items && items.All(x => x is JObject))
            {
                rows = new List<Dictionary<string, string>>();
                foreach (var it in items.Cast<JObject>())
                {
                    var dict = new Dictionary<string, string>();
                    foreach (var p in it.Properties())
                    {
                        var v = p.Value;
                        dict[p.Name] = v.Type == JTokenType.Null ? string.Empty : v.ToString(Formatting.None);
                    }
                    rows.Add(dict);
                }
            }
            else
            {
                // fallback sur aplat
                rows = BuildRows(data);
            }

            AppendLog($"Génération Excel: {xlsxPath}");
            return WriteExcel(xlsxPath, categoryName, rows);
        }

        private static List<Dictionary<string, string>> BuildRows(JToken token)
        {
            var rows = new List<Dictionary<string, string>>();

            if (token is JArray array)
            {
                if (array.Count == 0)
                {
                    rows.Add(new Dictionary<string, string>());
                }
                else if (array.All(x => x.Type == JTokenType.Object))
                {
                    foreach (var obj in array.Cast<JObject>())
                    {
                        var dict = new Dictionary<string, string>();
                        Flatten(obj, dict, prefix: null);
                        rows.Add(dict);
                    }
                }
                else
                {
                    // Mélange ou scalaires: une valeur par ligne sous colonne "value"
                    foreach (var x in array)
                    {
                        rows.Add(new Dictionary<string, string> { ["value"] = x.Type == JTokenType.Null ? string.Empty : x.ToString(Formatting.None) });
                    }
                }
            }
            else if (token is JObject o)
            {
                var dict = new Dictionary<string, string>();
                Flatten(o, dict, prefix: null);
                rows.Add(dict);
            }
            else
            {
                rows.Add(new Dictionary<string, string> { ["value"] = token.Type == JTokenType.Null ? string.Empty : token.ToString(Formatting.None) });
            }

            return rows;
        }

        private static void Flatten(JToken token, Dictionary<string, string> output, string? prefix)
        {
            switch (token.Type)
            {
                case JTokenType.Object:
                    foreach (var prop in ((JObject)token).Properties())
                    {
                        var key = string.IsNullOrEmpty(prefix) ? prop.Name : $"{prefix}.{prop.Name}";
                        Flatten(prop.Value, output, key);
                    }
                    break;
                case JTokenType.Array:
                    int index = 0;
                    foreach (var item in (JArray)token)
                    {
                        var key = string.IsNullOrEmpty(prefix) ? $"[{index}]" : $"{prefix}[{index}]";
                        Flatten(item, output, key);
                        index++;
                    }
                    break;
                default:
                    output[prefix ?? "value"] = token.Type == JTokenType.Null ? string.Empty : token.ToString(Formatting.None);
                    break;
            }
        }

        private void TryDeleteDirectoryContents(string dir)
        {
            try
            {
                foreach (var file in Directory.GetFiles(dir))
                {
                    try { File.Delete(file); } catch (Exception ex) { AppendLog($"Impossible de supprimer {Path.GetFileName(file)}: {ex.Message}"); }
                }
                foreach (var sub in Directory.GetDirectories(dir))
                {
                    try { Directory.Delete(sub, true); } catch (Exception ex) { AppendLog($"Impossible de supprimer le dossier {Path.GetFileName(sub)}: {ex.Message}"); }
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Erreur de suppression: {ex.Message}");
            }
        }

        private sealed class AppConfig
        {
            public string EndpointUrl { get; set; } = string.Empty;
            public string ApiKey { get; set; } = string.Empty;
            public string OutputFolder { get; set; } = string.Empty;
            public bool DeleteOld { get; set; }
        }
    }
}
