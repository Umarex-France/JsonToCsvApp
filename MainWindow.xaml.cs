using CsvHelper;
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

namespace JsonToCsvApp
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
            _apiVisible = false;
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
                        AppendLog("Aucune clé 'categories' trouvée, génération d’un CSV unique.");
                        string defaultName = "categorie";
                        await GenerateCsvForCategory(outDir, defaultName, root);
                    }
                    else if (categoriesToken is JArray arr)
                    {
                        int count = 0;
                        for (int i = 0; i < arr.Count; i++)
                        {
                            var cat = arr[i];
                            string catName = ResolveCategoryName(cat, i);
                            await GenerateCsvForCategory(outDir, catName, cat);
                            count++;
                        }
                        AppendLog($"{count} fichiers CSV générés dans {outDir}");
                    }
                    else
                    {
                        string catName = ResolveCategoryName(categoriesToken, 0);
                        await GenerateCsvForCategory(outDir, catName, categoriesToken);
                        AppendLog($"1 fichier CSV généré dans {outDir}");
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

            int totalCsv = 0;
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

                string fileName = ToCamelCase(famName) + ".csv";
                string csvPath = Path.Combine(familleFolder, fileName);
                await WriteCsv(csvPath, rows, allCaracColsGroup, nonEmptyCaracColsGroup);
                AppendLog($"CSV généré: {csvPath}");
                totalCsv++;
            }

            AppendLog($"{totalCsv} fichiers CSV générés dans {outDir}");

            // Génère l'index global nettoyé à la racine
            string indexPath = Path.Combine(outDir, "index.csv");
            await WriteCsv(indexPath, allRowsGlobal, allCaracColsGlobal, nonEmptyCaracColsGlobal);
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

                var val = prop.Value;
                switch (val.Type)
                {
                    case JTokenType.Array:
                        if (val is JArray arr)
                        {
                            // Joindre les tableaux en chaîne lisible
                            var joined = string.Join(" | ", arr.Select(x => x.Type == JTokenType.Null ? string.Empty : x.ToString(Formatting.None)));
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
                        dict[name] = val.ToString(Formatting.None);
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
                allCaracColumns?.Add(text);
                var valueToken = item["value"];
                if (valueToken == null || valueToken.Type == JTokenType.Null) continue;
                var value = valueToken.ToString(Formatting.None);
                if (string.IsNullOrWhiteSpace(value)) continue;
                dict[text] = value;
                nonEmptyCaracColumns?.Add(text);
            }
        }

        private static async Task WriteCsv(string path, List<Dictionary<string, string>> rows, ISet<string>? allCaracColumns = null, ISet<string>? nonEmptyCaracColumns = null)
        {
            await using var writer = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

            var headers = rows.SelectMany(d => d.Keys).Distinct().ToList();
            if (allCaracColumns != null && nonEmptyCaracColumns != null)
            {
                headers = headers
                    .Where(h => !allCaracColumns.Contains(h) || nonEmptyCaracColumns.Contains(h))
                    .ToList();
            }
            foreach (var h in headers) csv.WriteField(h);
            await csv.NextRecordAsync();

            foreach (var row in rows)
            {
                foreach (var h in headers)
                {
                    row.TryGetValue(h, out var val);
                    csv.WriteField(val ?? string.Empty);
                }
                await csv.NextRecordAsync();
            }
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

        private async Task GenerateCsvForCategory(string outputDir, string categoryName, JToken data)
        {
            string categoryFolder = Path.Combine(outputDir, categoryName);
            Directory.CreateDirectory(categoryFolder);
            string csvPath = Path.Combine(categoryFolder, $"{categoryName}.csv");

            var rows = BuildRows(data);
            AppendLog($"Génération CSV: {csvPath}");

            await using var writer = new StreamWriter(csvPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);

            var headers = rows.SelectMany(d => d.Keys).Distinct().ToList();
            foreach (var h in headers) csv.WriteField(h);
            await csv.NextRecordAsync();

            foreach (var row in rows)
            {
                foreach (var h in headers)
                {
                    row.TryGetValue(h, out var val);
                    csv.WriteField(val ?? string.Empty);
                }
                await csv.NextRecordAsync();
            }
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
