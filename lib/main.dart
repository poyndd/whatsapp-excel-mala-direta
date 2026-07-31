import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';

import 'package:excel/excel.dart' hide Border;
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:share_plus/share_plus.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:url_launcher/url_launcher.dart';

void main() {
  runApp(const WhatsExcelApp());
}

class WhatsExcelApp extends StatelessWidget {
  const WhatsExcelApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'WhatsApp Excel Mala Direta',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        colorSchemeSeed: Colors.green,
        useMaterial3: true,
      ),
      home: const HomePage(),
    );
  }
}

class MessageTemplate {
  final String name;
  final String message;
  final String? preferredPhoneColumn;
  final DateTime createdAt;
  final int version;

  MessageTemplate({
    required this.name,
    required this.message,
    this.preferredPhoneColumn,
    required this.createdAt,
    this.version = 1,
  });

  Map<String, dynamic> toJson() {
    return {
      'nome': name,
      'mensagem': message,
      'colunaTelefonePreferida': preferredPhoneColumn,
      'criadoEm': createdAt.toIso8601String(),
      'versao': version,
    };
  }

  factory MessageTemplate.fromJson(Map<String, dynamic> json) {
    return MessageTemplate(
      name: (json['nome'] ?? '').toString(),
      message: (json['mensagem'] ?? '').toString(),
      preferredPhoneColumn: json['colunaTelefonePreferida']?.toString(),
      createdAt: DateTime.tryParse((json['criadoEm'] ?? '').toString()) ??
          DateTime.now(),
      version: int.tryParse((json['versao'] ?? '1').toString()) ?? 1,
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  final TextEditingController templateNameController = TextEditingController();
  final TextEditingController messageController = TextEditingController();

  final List<String> logs = [];
  final List<MessageTemplate> savedTemplates = [];

  Uint8List? selectedExcelBytes;
  String? selectedFileName;
  String? selectedSheetName;
  String? selectedPhoneColumn;
  String? detectedPhoneColumn;

  List<String> sheetNames = [];
  List<String> headers = [];
  List<Map<String, String>> rows = [];

  int currentIndex = 0;
  int currentPhoneIndex = 0;
  bool isSending = false;

  static const String templatesPrefsKey = 'whatsapp_excel_templates_v1';

  @override
  void initState() {
    super.initState();
    loadTemplates();
  }

  @override
  void dispose() {
    templateNameController.dispose();
    messageController.dispose();
    super.dispose();
  }

  Future<void> loadTemplates() async {
    final prefs = await SharedPreferences.getInstance();
    final raw = prefs.getString(templatesPrefsKey);

    savedTemplates.clear();

    if (raw != null && raw.trim().isNotEmpty) {
      try {
        final decoded = jsonDecode(raw);
        if (decoded is List) {
          for (final item in decoded) {
            if (item is Map<String, dynamic>) {
              savedTemplates.add(MessageTemplate.fromJson(item));
            } else if (item is Map) {
              savedTemplates.add(
                MessageTemplate.fromJson(Map<String, dynamic>.from(item)),
              );
            }
          }
        }
      } catch (_) {
        addLog('Erro ao carregar modelos salvos.');
      }
    }

    setState(() {});
  }

  Future<void> persistTemplates() async {
    final prefs = await SharedPreferences.getInstance();
    final encoded = jsonEncode(
      savedTemplates.map((template) => template.toJson()).toList(),
    );
    await prefs.setString(templatesPrefsKey, encoded);
  }

  void addLog(String message) {
    final time = TimeOfDay.now();
    final prefix =
        '${time.hour.toString().padLeft(2, '0')}:${time.minute.toString().padLeft(2, '0')}';
    setState(() {
      logs.insert(0, '[$prefix] $message');
    });
  }

  Future<void> pickExcelFile() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
      withData: true,
    );

    if (result == null || result.files.isEmpty) {
      return;
    }

    final file = result.files.first;

    if (file.bytes == null) {
      addLog('Não foi possível ler o arquivo selecionado.');
      return;
    }

    selectedExcelBytes = file.bytes;
    selectedFileName = file.name;

    try {
      final excel = Excel.decodeBytes(selectedExcelBytes!);
      sheetNames = excel.tables.keys.toList();

      if (sheetNames.isEmpty) {
        addLog('O Excel não possui planilhas válidas.');
        return;
      }

      selectedSheetName = sheetNames.first;
      parseSelectedSheet();

      addLog('Excel carregado: ${file.name}');
    } catch (e) {
      addLog('Erro ao abrir Excel: $e');
    }

    setState(() {});
  }

  void parseSelectedSheet() {
    if (selectedExcelBytes == null || selectedSheetName == null) {
      return;
    }

    headers.clear();
    rows.clear();
    selectedPhoneColumn = null;
    detectedPhoneColumn = null;
    currentIndex = 0;
    currentPhoneIndex = 0;

    final excel = Excel.decodeBytes(selectedExcelBytes!);
    final sheet = excel.tables[selectedSheetName];

    if (sheet == null || sheet.rows.isEmpty) {
      addLog('Planilha vazia ou inválida.');
      setState(() {});
      return;
    }

    final firstRow = sheet.rows.first;

    headers = firstRow
        .map((cell) => (cell?.value?.toString() ?? '').trim())
        .where((value) => value.isNotEmpty)
        .toList();

    if (headers.isEmpty) {
      addLog('Não encontrei cabeçalhos na primeira linha.');
      setState(() {});
      return;
    }

    for (int r = 1; r < sheet.rows.length; r++) {
      final row = sheet.rows[r];
      final Map<String, String> mapped = {};

      for (int c = 0; c < headers.length; c++) {
        final header = headers[c];
        final value = c < row.length ? row[c]?.value?.toString() ?? '' : '';
        mapped[header] = value.trim();
      }

      final hasAnyValue = mapped.values.any((value) => value.trim().isNotEmpty);
      if (hasAnyValue) {
        rows.add(mapped);
      }
    }

    detectedPhoneColumn = detectPhoneColumn(headers, rows);
    selectedPhoneColumn = detectedPhoneColumn ?? headers.first;

    addLog('Planilha "$selectedSheetName" carregada com ${rows.length} linhas.');
    if (detectedPhoneColumn != null) {
      addLog('Coluna de telefone sugerida: $detectedPhoneColumn');
    } else {
      addLog('Não foi possível sugerir a coluna de telefone.');
    }

    setState(() {});
  }

  String? detectPhoneColumn(
    List<String> headers,
    List<Map<String, String>> rows,
  ) {
    if (headers.isEmpty) return null;

    final Map<String, int> scores = {};

    for (final header in headers) {
      final normalized = normalizeText(header);
      int score = 0;

      if (normalized.contains('celular')) score += 120;
      if (normalized.contains('whatsapp')) score += 120;
      if (normalized.contains('zap')) score += 110;
      if (normalized.contains('telefone')) score += 100;
      if (normalized.contains('fone')) score += 90;
      if (normalized == 'tel') score += 80;
      if (normalized.contains('tel ')) score += 60;
      if (normalized.contains('mobile')) score += 80;
      if (normalized.contains('phone')) score += 80;
      if (normalized.contains('contato')) score += 40;
      if (normalized.contains('numero')) score += 30;
      if (normalized.contains('nro')) score += 25;

      int checked = 0;
      int validPhones = 0;

      for (final row in rows.take(80)) {
        final value = row[header] ?? '';
        if (value.trim().isEmpty) continue;

        checked++;
        final phones = extractPhonesFromText(value);

        if (phones.isNotEmpty) {
          validPhones++;
        }
      }

      if (checked > 0) {
        final rate = validPhones / checked;

        if (rate >= 0.80) score += 100;
        if (rate >= 0.70) score += 80;
        if (rate >= 0.50) score += 50;
        if (rate >= 0.30) score += 25;
      }

      scores[header] = score;
    }

    final sorted = scores.entries.toList()
      ..sort((a, b) => b.value.compareTo(a.value));

    if (sorted.isEmpty || sorted.first.value <= 0) {
      return null;
    }

    return sorted.first.key;
  }

  String normalizeText(String value) {
    return value
        .toLowerCase()
        .replaceAll('á', 'a')
        .replaceAll('à', 'a')
        .replaceAll('ã', 'a')
        .replaceAll('â', 'a')
        .replaceAll('é', 'e')
        .replaceAll('ê', 'e')
        .replaceAll('í', 'i')
        .replaceAll('ó', 'o')
        .replaceAll('õ', 'o')
        .replaceAll('ô', 'o')
        .replaceAll('ú', 'u')
        .replaceAll('ü', 'u')
        .replaceAll('ç', 'c')
        .trim();
  }

  String onlyDigits(String value) {
    return value.replaceAll(RegExp(r'[^0-9]'), '');
  }

  String formatPhoneForWhatsApp(String rawPhone) {
    String digits = onlyDigits(rawPhone);

    if (digits.startsWith('00')) {
      digits = digits.substring(2);
    }

    if (!digits.startsWith('55')) {
      digits = '55$digits';
    }

    return digits;
  }

  bool isValidPhoneForWhatsApp(String rawPhone) {
    final digits = formatPhoneForWhatsApp(rawPhone);
    return digits.length >= 12 && digits.length <= 13;
  }

  bool isPhoneLikeHeader(String header) {
    final normalized = normalizeText(header);

    return normalized.contains('celular') ||
        normalized.contains('cel') ||
        normalized.contains('whatsapp') ||
        normalized.contains('zap') ||
        normalized.contains('telefone') ||
        normalized.contains('telef') ||
        normalized.contains('fone') ||
        normalized == 'tel' ||
        normalized.startsWith('tel ') ||
        normalized.contains(' tel') ||
        normalized.contains('mobile') ||
        normalized.contains('phone') ||
        normalized.contains('contato');
  }

  List<String> extractPhonesFromText(String rawText) {
    final Set<String> phones = {};

    if (rawText.trim().isEmpty) {
      return [];
    }

    final text = rawText
        .replaceAll('\n', ' ')
        .replaceAll('\r', ' ')
        .replaceAll(';', ' ')
        .replaceAll(',', ' ')
        .replaceAll('|', ' ')
        .replaceAll('/', ' ')
        .replaceAll('\\', ' ');

    final phonePattern = RegExp(
      r'(\+?55\s*)?(\(?\d{2}\)?\s*)?(9?\d{4}[-\s]?\d{4})',
    );

    for (final match in phonePattern.allMatches(text)) {
      final candidate = match.group(0) ?? '';
      final digits = onlyDigits(candidate);

      if (digits.length >= 10 && digits.length <= 13) {
        final formatted = formatPhoneForWhatsApp(candidate);

        if (formatted.length >= 12 && formatted.length <= 13) {
          phones.add(formatted);
        }
      }
    }

    final fullDigits = onlyDigits(rawText);

    if (fullDigits.length >= 10 && fullDigits.length <= 13) {
      final formatted = formatPhoneForWhatsApp(rawText);

      if (formatted.length >= 12 && formatted.length <= 13) {
        phones.add(formatted);
      }
    }

    return phones.toList();
  }

  List<String> getValidPhonesForRow(Map<String, String> row) {
    final Set<String> phones = {};

    void addPhonesFromHeader(String header) {
      final value = row[header] ?? '';

      for (final phone in extractPhonesFromText(value)) {
        if (phone.length >= 12 && phone.length <= 13) {
          phones.add(phone);
        }
      }
    }

    if (selectedPhoneColumn != null && row.containsKey(selectedPhoneColumn)) {
      addPhonesFromHeader(selectedPhoneColumn!);
    }

    for (final header in headers) {
      if (header == selectedPhoneColumn) {
        continue;
      }

      if (isPhoneLikeHeader(header)) {
        addPhonesFromHeader(header);
      }
    }

    return phones.toList();
  }

  String mergeMessage(String template, Map<String, String> row) {
    String result = template;

    for (final entry in row.entries) {
      result = result.replaceAll('{${entry.key}}', entry.value);
    }

    return result;
  }

  void insertMergeField(String header) {
    final text = messageController.text;
    final selection = messageController.selection;
    final field = '{$header}';

    final start = selection.start < 0 ? text.length : selection.start;
    final end = selection.end < 0 ? text.length : selection.end;

    final newText = text.replaceRange(start, end, field);

    messageController.text = newText;
    messageController.selection = TextSelection.collapsed(
      offset: start + field.length,
    );

    setState(() {});
  }

  String getCurrentPreviewMessage() {
    if (rows.isEmpty) return '';
    if (currentIndex < 0 || currentIndex >= rows.length) return '';

    return mergeMessage(
      messageController.text,
      rows[currentIndex],
    );
  }

  List<String> getCurrentPhones() {
    if (rows.isEmpty) return [];
    if (currentIndex < 0 || currentIndex >= rows.length) return [];

    return getValidPhonesForRow(rows[currentIndex]);
  }

  String getCurrentPhoneRaw() {
    final phones = getCurrentPhones();

    if (phones.isEmpty) {
      return '';
    }

    final safeIndex =
        currentPhoneIndex >= phones.length ? phones.length - 1 : currentPhoneIndex;

    return phones[safeIndex];
  }

  Future<void> openWhatsAppForCurrentRow() async {
    if (rows.isEmpty) {
      addLog('Não há linhas carregadas.');
      return;
    }

    if (messageController.text.trim().isEmpty) {
      addLog('Digite uma mensagem antes de enviar.');
      return;
    }

    final row = rows[currentIndex];
    final phones = getValidPhonesForRow(row);

    if (phones.isEmpty) {
      addLog('Nenhum celular válido encontrado na linha ${currentIndex + 2}.');
      return;
    }

    if (currentPhoneIndex >= phones.length) {
      currentPhoneIndex = 0;
    }

    final phone = phones[currentPhoneIndex];
    final message = mergeMessage(messageController.text, row);
    final encodedMessage = Uri.encodeComponent(message);

    final appUri = Uri.parse(
      'whatsapp://send?phone=$phone&text=$encodedMessage',
    );

    final webUri = Uri.parse(
      'https://wa.me/$phone?text=$encodedMessage',
    );

    try {
      bool opened = false;

      if (await canLaunchUrl(appUri)) {
        opened = await launchUrl(
          appUri,
          mode: LaunchMode.externalApplication,
        );
      }

      if (!opened && await canLaunchUrl(webUri)) {
        opened = await launchUrl(
          webUri,
          mode: LaunchMode.externalApplication,
        );
      }

      if (opened) {
        addLog(
          'WhatsApp aberto para linha ${currentIndex + 2}, celular ${currentPhoneIndex + 1}/${phones.length}: $phone',
        );

        setState(() {
          if (currentPhoneIndex < phones.length - 1) {
            currentPhoneIndex++;
          } else {
            currentPhoneIndex = 0;

            if (currentIndex < rows.length - 1) {
              currentIndex++;
              addLog('Todos os celulares da linha anterior foram abertos. Próxima linha preparada.');
            } else {
              addLog('Todos os celulares da última linha foram abertos.');
            }
          }
        });
      } else {
        addLog('Não foi possível abrir o WhatsApp.');
      }
    } catch (e) {
      addLog('Erro ao abrir WhatsApp: $e');
    }
  }

  void goToNextRow() {
    if (rows.isEmpty) return;

    final phones = getValidPhonesForRow(rows[currentIndex]);

    if (phones.isNotEmpty && currentPhoneIndex < phones.length - 1) {
      setState(() {
        currentPhoneIndex++;
      });

      addLog(
        'Avançou para o celular ${currentPhoneIndex + 1}/${phones.length} da linha ${currentIndex + 2}.',
      );

      return;
    }

    if (currentIndex < rows.length - 1) {
      setState(() {
        currentIndex++;
        currentPhoneIndex = 0;
      });

      addLog('Avançou para a linha ${currentIndex + 2}.');
    } else {
      addLog('Fim da base.');
    }
  }

  void goToPreviousRow() {
    if (rows.isEmpty) return;

    if (currentPhoneIndex > 0) {
      setState(() {
        currentPhoneIndex--;
      });

      final phones = getValidPhonesForRow(rows[currentIndex]);

      addLog(
        'Voltou para o celular ${currentPhoneIndex + 1}/${phones.length} da linha ${currentIndex + 2}.',
      );

      return;
    }

    if (currentIndex > 0) {
      setState(() {
        currentIndex--;
        final previousPhones = getValidPhonesForRow(rows[currentIndex]);
        currentPhoneIndex = previousPhones.isEmpty ? 0 : previousPhones.length - 1;
      });

      addLog('Voltou para a linha ${currentIndex + 2}.');
    }
  }

  Future<void> saveCurrentTemplate() async {
    final name = templateNameController.text.trim();
    final message = messageController.text;

    if (name.isEmpty) {
      addLog('Informe o nome do modelo.');
      return;
    }

    if (message.trim().isEmpty) {
      addLog('Digite a mensagem do modelo.');
      return;
    }

    final template = MessageTemplate(
      name: name,
      message: message,
      preferredPhoneColumn: selectedPhoneColumn,
      createdAt: DateTime.now(),
    );

    final existingIndex = savedTemplates.indexWhere(
      (item) => normalizeText(item.name) == normalizeText(name),
    );

    if (existingIndex >= 0) {
      savedTemplates[existingIndex] = template;
      addLog('Modelo atualizado: $name');
    } else {
      savedTemplates.add(template);
      addLog('Modelo salvo: $name');
    }

    await persistTemplates();
    setState(() {});
  }

  void loadTemplateIntoEditor(MessageTemplate template) {
    templateNameController.text = template.name;
    messageController.text = template.message;

    if (template.preferredPhoneColumn != null &&
        headers.contains(template.preferredPhoneColumn)) {
      selectedPhoneColumn = template.preferredPhoneColumn;
      currentPhoneIndex = 0;
    }

    addLog('Modelo carregado: ${template.name}');
    setState(() {});
  }

  Future<void> importTemplateJson() async {
    final result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['json'],
      withData: true,
    );

    if (result == null || result.files.isEmpty) {
      return;
    }

    final file = result.files.first;

    if (file.bytes == null) {
      addLog('Não foi possível ler o JSON.');
      return;
    }

    try {
      final raw = utf8.decode(file.bytes!);
      final decoded = jsonDecode(raw);

      if (decoded is Map) {
        final template = MessageTemplate.fromJson(
          Map<String, dynamic>.from(decoded),
        );

        final existingIndex = savedTemplates.indexWhere(
          (item) => normalizeText(item.name) == normalizeText(template.name),
        );

        if (existingIndex >= 0) {
          savedTemplates[existingIndex] = template;
        } else {
          savedTemplates.add(template);
        }

        await persistTemplates();
        loadTemplateIntoEditor(template);
        addLog('Modelo importado: ${template.name}');
      } else {
        addLog('JSON inválido. O arquivo precisa conter um único modelo.');
      }
    } catch (e) {
      addLog('Erro ao importar modelo: $e');
    }

    setState(() {});
  }

  Future<void> exportCurrentTemplateJson() async {
    final name = templateNameController.text.trim();
    final message = messageController.text;

    if (name.isEmpty || message.trim().isEmpty) {
      addLog('Informe nome e mensagem para exportar.');
      return;
    }

    final template = MessageTemplate(
      name: name,
      message: message,
      preferredPhoneColumn: selectedPhoneColumn,
      createdAt: DateTime.now(),
    );

    try {
      final dir = await getTemporaryDirectory();
      final safeName = name
          .replaceAll(RegExp(r'[^a-zA-Z0-9_\-]'), '_')
          .replaceAll('__', '_');

      final file = File('${dir.path}/$safeName.json');

      await file.writeAsString(
        const JsonEncoder.withIndent('  ').convert(template.toJson()),
        encoding: utf8,
      );

      await Share.shareXFiles(
        [XFile(file.path)],
        text: 'Modelo de mensagem WhatsApp Excel',
      );

      addLog('Modelo exportado: $safeName.json');
    } catch (e) {
      addLog('Erro ao exportar modelo: $e');
    }
  }

  Future<void> deleteTemplate(MessageTemplate template) async {
    savedTemplates.removeWhere(
      (item) => normalizeText(item.name) == normalizeText(template.name),
    );

    await persistTemplates();
    addLog('Modelo excluído: ${template.name}');
    setState(() {});
  }

  Widget buildExcelSection() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '1. Arquivo Excel',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18),
            ),
            const SizedBox(height: 8),
            FilledButton.icon(
              onPressed: pickExcelFile,
              icon: const Icon(Icons.upload_file),
              label: const Text('Selecionar Excel .xlsx'),
            ),
            const SizedBox(height: 8),
            Text(selectedFileName == null
                ? 'Nenhum arquivo selecionado.'
                : 'Arquivo: $selectedFileName'),
            const SizedBox(height: 8),
            if (sheetNames.isNotEmpty)
              DropdownButtonFormField<String>(
                value: selectedSheetName,
                decoration: const InputDecoration(
                  labelText: 'Planilha',
                  border: OutlineInputBorder(),
                ),
                items: sheetNames
                    .map(
                      (sheet) => DropdownMenuItem(
                        value: sheet,
                        child: Text(sheet),
                      ),
                    )
                    .toList(),
                onChanged: (value) {
                  if (value == null) return;
                  selectedSheetName = value;
                  parseSelectedSheet();
                },
              ),
            const SizedBox(height: 8),
            if (headers.isNotEmpty)
              DropdownButtonFormField<String>(
                value: selectedPhoneColumn,
                decoration: InputDecoration(
                  labelText: detectedPhoneColumn == null
                      ? 'Coluna de telefone principal'
                      : 'Coluna de telefone sugerida: $detectedPhoneColumn',
                  border: const OutlineInputBorder(),
                ),
                items: headers
                    .map(
                      (header) => DropdownMenuItem(
                        value: header,
                        child: Text(header),
                      ),
                    )
                    .toList(),
                onChanged: (value) {
                  setState(() {
                    selectedPhoneColumn = value;
                    currentPhoneIndex = 0;
                  });
                },
              ),
            const SizedBox(height: 8),
            Text('Registros carregados: ${rows.length}'),
            const SizedBox(height: 4),
            const Text(
              'Obs.: além da coluna principal, o app também procura celulares válidos em outras colunas com nomes parecidos com celular, telefone, WhatsApp, zap, fone ou contato.',
              style: TextStyle(fontSize: 12),
            ),
          ],
        ),
      ),
    );
  }

  Widget buildTemplateSection() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '2. Modelos de mensagem',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18),
            ),
            const SizedBox(height: 8),
            if (savedTemplates.isNotEmpty)
              DropdownButtonFormField<String>(
                decoration: const InputDecoration(
                  labelText: 'Carregar modelo salvo',
                  border: OutlineInputBorder(),
                ),
                items: savedTemplates
                    .map(
                      (template) => DropdownMenuItem(
                        value: template.name,
                        child: Text(template.name),
                      ),
                    )
                    .toList(),
                onChanged: (value) {
                  if (value == null) return;

                  final template = savedTemplates.firstWhere(
                    (item) => item.name == value,
                  );

                  loadTemplateIntoEditor(template);
                },
              ),
            const SizedBox(height: 8),
            TextField(
              controller: templateNameController,
              decoration: const InputDecoration(
                labelText: 'Nome do modelo',
                hintText: 'Ex: Campanha regularização',
                border: OutlineInputBorder(),
              ),
            ),
            const SizedBox(height: 8),
            Row(
              children: [
                Expanded(
                  child: FilledButton.icon(
                    onPressed: saveCurrentTemplate,
                    icon: const Icon(Icons.save),
                    label: const Text('Salvar'),
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: importTemplateJson,
                    icon: const Icon(Icons.file_open),
                    label: const Text('Importar'),
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: exportCurrentTemplateJson,
                    icon: const Icon(Icons.ios_share),
                    label: const Text('Exportar'),
                  ),
                ),
              ],
            ),
            const SizedBox(height: 8),
            if (savedTemplates.isNotEmpty)
              Wrap(
                spacing: 8,
                runSpacing: 8,
                children: savedTemplates.map((template) {
                  return InputChip(
                    label: Text(template.name),
                    onPressed: () => loadTemplateIntoEditor(template),
                    onDeleted: () => deleteTemplate(template),
                  );
                }).toList(),
              ),
          ],
        ),
      ),
    );
  }

  Widget buildEditorSection() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '3. Editor da mensagem',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18),
            ),
            const SizedBox(height: 8),
            TextField(
              controller: messageController,
              minLines: 8,
              maxLines: 14,
              decoration: const InputDecoration(
                labelText: 'Mensagem',
                hintText:
                    'Ex: Olá {Nome}, temos uma condição especial para o contrato {Contrato}.',
                border: OutlineInputBorder(),
              ),
              onChanged: (_) => setState(() {}),
            ),
            const SizedBox(height: 8),
            const Text(
              'Campos de mesclagem',
              style: TextStyle(fontWeight: FontWeight.w600),
            ),
            const SizedBox(height: 8),
            if (headers.isEmpty)
              const Text('Carregue um Excel para exibir os campos.')
            else
              Wrap(
                spacing: 8,
                runSpacing: 8,
                children: headers.map((header) {
                  return ActionChip(
                    label: Text(header),
                    onPressed: () => insertMergeField(header),
                  );
                }).toList(),
              ),
          ],
        ),
      ),
    );
  }

  Widget buildPreviewSection() {
    final preview = getCurrentPreviewMessage();
    final phones = getCurrentPhones();

    final safePhoneIndex = phones.isEmpty
        ? 0
        : currentPhoneIndex >= phones.length
            ? phones.length - 1
            : currentPhoneIndex;

    final currentPhone = phones.isEmpty ? '' : phones[safePhoneIndex];

    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '4. Prévia e envio',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18),
            ),
            const SizedBox(height: 8),
            if (rows.isEmpty)
              const Text('Nenhuma linha carregada.')
            else ...[
              Text('Linha atual no Excel: ${currentIndex + 2}'),
              Text('Registro: ${currentIndex + 1} de ${rows.length}'),
              const SizedBox(height: 8),
              Text('Celulares válidos encontrados nesta linha: ${phones.length}'),
              Text(
                phones.isEmpty
                    ? 'Celular atual: nenhum'
                    : 'Celular atual: ${safePhoneIndex + 1}/${phones.length} - $currentPhone',
              ),
              const SizedBox(height: 8),
              Container(
                width: double.infinity,
                padding: const EdgeInsets.all(12),
                decoration: BoxDecoration(
                  border: Border.all(color: Colors.black26),
                  borderRadius: BorderRadius.circular(8),
                ),
                child: Text(
                  preview.isEmpty ? 'Prévia vazia.' : preview,
                ),
              ),
              const SizedBox(height: 8),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: goToPreviousRow,
                      icon: const Icon(Icons.arrow_back),
                      label: const Text('Anterior'),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: FilledButton.icon(
                      onPressed: openWhatsAppForCurrentRow,
                      icon: const Icon(Icons.send),
                      label: const Text('Abrir WhatsApp'),
                    ),
                  ),
                  const SizedBox(width: 8),
                  Expanded(
                    child: OutlinedButton.icon(
                      onPressed: goToNextRow,
                      icon: const Icon(Icons.arrow_forward),
                      label: const Text('Próximo'),
                    ),
                  ),
                ],
              ),
              const SizedBox(height: 8),
              const Text(
                'Observação: se a linha tiver mais de um celular válido, o app prepara a mesma mensagem para cada celular, um por vez.',
                style: TextStyle(fontSize: 12),
              ),
            ],
          ],
        ),
      ),
    );
  }

  Widget buildLogSection() {
    return Card(
      child: Padding(
        padding: const EdgeInsets.all(12),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              '5. Log',
              style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18),
            ),
            const SizedBox(height: 8),
            if (logs.isEmpty)
              const Text('Nenhum evento registrado.')
            else
              SizedBox(
                height: 180,
                child: ListView.builder(
                  itemCount: logs.length,
                  itemBuilder: (context, index) {
                    return Text(
                      logs[index],
                      style: const TextStyle(fontSize: 12),
                    );
                  },
                ),
              ),
          ],
        ),
      ),
    );
  }

  Widget buildLgpdWarning() {
    return Card(
      color: Colors.amber.shade50,
      child: const Padding(
        padding: EdgeInsets.all(12),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Icon(Icons.warning_amber, color: Colors.orange),
            SizedBox(width: 8),
            Expanded(
              child: Text(
                'Atenção LGPD/CDC: revise cuidadosamente os campos de mesclagem antes do envio. Evite expor CPF, contrato, valores, situação de atraso ou outros dados sensíveis quando a comunicação puder ser vista por terceiros.',
              ),
            ),
          ],
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    final bottomPadding = MediaQuery.of(context).padding.bottom;

    return Scaffold(
      appBar: AppBar(
        title: const Text('WhatsApp Excel Mala Direta'),
        actions: [
          IconButton(
            tooltip: 'Limpar log',
            onPressed: () {
              setState(() {
                logs.clear();
              });
            },
            icon: const Icon(Icons.delete_sweep),
          ),
        ],
      ),
      body: SafeArea(
        child: SingleChildScrollView(
          padding: EdgeInsets.fromLTRB(12, 12, 12, bottomPadding + 12),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              buildLgpdWarning(),
              buildExcelSection(),
              buildTemplateSection(),
              buildEditorSection(),
              buildPreviewSection(),
              buildLogSection(),
            ],
          ),
        ),
      ),
    );
  }
}
