import 'dart:io';
import 'package:flutter/material.dart';
import 'package:flutter/painting.dart' as painting; // حل تعارض حدود التصميم مع مكتبة الإكسيل
import 'package:file_picker/file_picker.dart';
import 'package:excel/excel.dart' as px;
import 'package:mobile_scanner/mobile_scanner.dart';
import 'package:google_mlkit_text_recognition/google_mlkit_text_recognition.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const StugraScanApp());
}

class StugraScanApp extends StatelessWidget {
  const StugraScanApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'StugraScan',
      debugShowCheckedModeBanner: false,
      theme: ThemeData.dark(),
      home: const MainScreen(),
    );
  }
}

class MainScreen extends StatefulWidget {
  const MainScreen({super.key});

  @override
  State<MainScreen> createState() => _MainScreenState();
}

class _MainScreenState extends State<MainScreen> {
  String _fileName = "لم يتم اختيار ملف الكنترول بعد";
  String? _selectedFilePath;
  List<String> _subjects = []; 
  String? _selectedSubject;
  bool _isLoading = false;

  String _secretIdResult = "سيظهر هنا الرقم السري";
  final TextEditingController _gradeController = TextEditingController();

  int _totalStudents = 0;
  int _gradedStudents = 0;
  
  final MobileScannerController _cameraController = MobileScannerController(
    autoStart: false,
    torchEnabled: false,
  );
  bool _isScanningActive = false;
  bool _isTorchOn = false;

  bool _isLocalSaveEnabled = true;       
  bool _isOcrEnabled = true;             
  bool _isIndicEnhanceEnabled = true;    
  bool _isRedPenEnhanceEnabled = false;   

  final Map<String, String> _tempGradesCache = {};
  final TextRecognizer _textRecognizer = TextRecognizer(script: TextRecognitionScript.latin);

  String _convertArabicIndicToEnglish(String input) {
    const hindiDigits = ['٠', '١', '٢', '٣', '٤', '٥', '٦', '٧', '٨', '٩'];
    const englishDigits = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
    
    String result = input;
    for (int i = 0; i < hindiDigits.length; i++) {
      result = result.replaceAll(hindiDigits[i], englishDigits[i]);
    }
    return result;
  }

  String _extractDigits(String input) {
    final cleaned = _convertArabicIndicToEnglish(input);
    return cleaned.replaceAll(RegExp(r'[^0-9]'), '');
  }

  Future<void> _pickAndParseExcel() async {
    setState(() {
      _isLoading = true;
      _subjects.clear();
      _selectedSubject = null;
      _fileName = "جاري قراءة الملف...";
    });

    try {
      FilePickerResult? result = await FilePicker.platform.pickFiles(
        type: FileType.custom,
        allowedExtensions: ['xlsx', 'xls'],
      );

      if (result != null && result.files.single.path != null) {
        _selectedFilePath = result.files.single.path;
        String nameOfFile = result.files.single.name;
        
        var bytes = File(_selectedFilePath!).readAsBytesSync();
        var excel = px.Excel.decodeBytes(bytes);
        
        String firstSheet = excel.tables.keys.first;
        var sheet = excel.tables[firstSheet];

        // تم إصلاح السطر أدناه من maxCols إلى maxColumns لحل الخطأ التجميعي
        if (sheet != null && sheet.maxColumns > 0) {
          var firstRow = sheet.rows.first; 
          List<String> tempSubjects = [];

          int startColumn = 4;  
          int endColumn = 18;  

          for (int i = startColumn; i <= endColumn; i++) {
            if (i < firstRow.length) {
              var cellValue = firstRow[i]?.value;
              if (cellValue != null) {
                String subjectName = cellValue.toString().trim();
                if (subjectName.isNotEmpty) {
                  tempSubjects.add(subjectName);
                }
              }
            }
          }

          setState(() {
            _fileName = nameOfFile;
            _subjects = tempSubjects;
            _totalStudents = sheet.maxRows > 1 ? sheet.maxRows - 1 : 0; 
            _gradedStudents = 0;
            _tempGradesCache.clear(); 
          });

          if (_subjects.isEmpty) {
            _showSnackBar("تنبيه: لم يتم العثور على مواد في الأعمدة من E إلى S.");
          }
        }
      } else {
        setState(() {
          _fileName = _selectedFilePath != null ? _selectedFilePath!.split('/').last : "لم يتم اختيار ملف الكنترول بعد";
        });
      }
    } catch (e) {
      setState(() {
        _fileName = "فشل في قراءة ملف الأكسيل";
      });
      _showSnackBar("حدث خطأ أثناء المعالجة: $e");
    } finally {
      setState(() {
        _isLoading = false;
      });
    }
  }

  Future<void> _saveAndExportToExcel() async {
    if (_selectedFilePath == null || _selectedSubject == null || _tempGradesCache.isEmpty) {
      _showSnackBar("لا توجد بيانات ممسوحة جديدة لحفظها حالياً!");
      return;
    }

    setState(() {
      _isLoading = true;
    });

    try {
      var bytes = File(_selectedFilePath!).readAsBytesSync();
      var excel = px.Excel.decodeBytes(bytes);
      String firstSheet = excel.tables.keys.first;
      var sheet = excel.tables[firstSheet];

      if (sheet != null) {
        var firstRow = sheet.rows.first;
        int targetColumnIndex = -1;

        for (int i = 0; i < firstRow.length; i++) {
          if (firstRow[i]?.value?.toString().trim() == _selectedSubject) {
            targetColumnIndex = i;
            break;
          }
        }

        if (targetColumnIndex == -1) {
          _showSnackBar("خطأ: لم يتم العثور على عمود المادة في الملف!");
          return;
        }

        int updatedCount = 0;
        for (int rowIdx = 1; rowIdx < sheet.maxRows; rowIdx++) {
          var row = sheet.rows[rowIdx];
          var secretIdCell = row.length > 2 ? row[2]?.value?.toString().trim() : row[1]?.value?.toString().trim();

          if (secretIdCell != null && _tempGradesCache.containsKey(secretIdCell)) {
            String gradeToPut = _tempGradesCache[secretIdCell]!;
            
            sheet.updateCell(
              px.CellIndex.indexByColumnRow(columnIndex: targetColumnIndex, rowIndex: rowIdx),
              px.CellValue.withValue(gradeToPut),
            );
            updatedCount++;
          }
        }

        var fileBytes = excel.save();
        if (fileBytes != null) {
          File(_selectedFilePath!)
            ..createSync(recursive: true)
            ..writeAsBytesSync(fileBytes);
          
          _showSnackBar("🎉 تم حفظ وتعديل الملف بنجاح! تم رصد $updatedCount طلاب.");
          setState(() {
            _tempGradesCache.clear(); 
          });
        }
      }
    } catch (e) {
      _showSnackBar("حدث خطأ أثناء حفظ الملف: $e");
    } finally {
      setState(() {
        _isLoading = false;
      });
    }
  }

  void _toggleScanning() {
    if (_selectedSubject == null) {
      _showSnackBar("يرجى اختيار المادة المراد رصدها أولاً قبل بدء المسح!");
      return;
    }

    setState(() {
      _isScanningActive = !_isScanningActive;
    });

    if (_isScanningActive) {
      _cameraController.start();
    } else {
      _cameraController.stop();
    }
  }

  Future<void> _processCapturedImage(BarcodeCapture capture) async {
    final List<Barcode> barcodes = capture.barcodes;
    
    if (barcodes.isNotEmpty && barcodes.first.rawValue != null) {
      final String qrValue = barcodes.first.rawValue!;
      _cameraController.stop();

      String detectedSubjectCode = "";
      String detectedGrade = "";

      if (_isOcrEnabled && capture.image != null) {
        final InputImage inputImage = InputImage.fromBytes(
          bytes: capture.image!,
          metadata: InputImageMetadata(
            size: Size(capture.width!.toDouble(), capture.height!.toDouble()),
            rotation: InputImageRotation.rotation0, 
            format: InputImageFormat.nv21, 
            bytesPerRow: capture.width!,
          ),
        );

        try {
          final RecognizedText recognizedText = await _textRecognizer.processImage(inputImage);
          
          for (TextBlock block in recognizedText.blocks) {
            for (TextLine line in block.lines) {
              String lineText = line.text.trim();
              
              if (lineText.contains("المادة") || lineText.contains("كود") || (lineText.length == 3 && RegExp(r'^\d+$').hasMatch(_extractDigits(lineText)))) {
                detectedSubjectCode = _extractDigits(lineText);
              } else {
                String numbersOnly = _isIndicEnhanceEnabled ? _extractDigits(lineText) : lineText.replaceAll(RegExp(r'[^0-9]'), '');
                if (numbersOnly.isNotEmpty && numbersOnly.length <= 3) {
                  detectedGrade = numbersOnly;
                }
              }
            }
          }

          int currentSubjectIndex = _subjects.indexOf(_selectedSubject!) + 1;

          if (detectedSubjectCode.isNotEmpty && detectedSubjectCode != currentSubjectIndex.toString()) {
            _showSnackBar("⚠️ كود المادة المقروء ($detectedSubjectCode) لا يطابق المادة المختارة!");
            _cameraController.start();
            return;
          }

        } catch (e) {
          debugPrint("خطأ أثناء قراءة الـ OCR: $e");
        }
      }

      setState(() {
        _secretIdResult = qrValue;
        if (detectedGrade.isNotEmpty) {
          _gradeController.text = detectedGrade;
        }
        
        if (_isLocalSaveEnabled) {
          if (!_tempGradesCache.containsKey(qrValue)) {
            _tempGradesCache[qrValue] = _gradeController.text;
            _gradedStudents = _tempGradesCache.length; 
          }
        } else {
          _gradedStudents += 1;
        }
        
        _isScanningActive = false;
      });

      _showSnackBar("تم مسح الكود بنجاح وحفظه مؤقتاً 🗳️");
    }
  }

  void _showSettingsDialog() {
    showDialog(
      context: context,
      builder: (BuildContext context) {
        return StatefulBuilder(
          builder: (context, setDialogState) {
            return AlertDialog(
              backgroundColor: const Color(0xFF121212),
              title: const Text(
                "خيارات الضبط",
                style: TextStyle(color: Colors.white),
                textAlign: TextAlign.center,
              ),
              content: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  const Align(
                    alignment: Alignment.centerRight,
                    child: Text("إعدادات الحفظ", style: TextStyle(color: Colors.green, fontWeight: FontWeight.bold)),
                  ),
                  SwitchListTile(
                    title: const Text("الحفظ المؤقت أولاً ثم التصدير", style: TextStyle(fontSize: 14)),
                    subtitle: const Text("سيتم التعديل مباشرة على الملف الأصلي المختار", style: TextStyle(fontSize: 11, color: Colors.grey)),
                    secondary: const Icon(Icons.save_as, color: Colors.white70),
                    value: _isLocalSaveEnabled,
                    activeColor: Colors.green,
                    onChanged: (bool value) {
                      setDialogState(() => _isLocalSaveEnabled = value);
                      setState(() => _isLocalSaveEnabled = value);
                    },
                  ),
                  const Divider(color: Colors.grey),
                  SwitchListTile(
                    title: const Text("التعرف على الكتابة اليدوية OCR", style: TextStyle(fontSize: 14)),
                    subtitle: const Text("سيتم رصد الدرجة المكتوبة بجانب الكود تلقائياً", style: TextStyle(fontSize: 11, color: Colors.grey)),
                    secondary: const Icon(Icons.text_fields, color: Colors.white70),
                    value: _isOcrEnabled,
                    activeColor: Colors.green,
                    onChanged: (bool value) {
                      setDialogState(() => _isOcrEnabled = value);
                      setState(() => _isOcrEnabled = value);
                    },
                  ),
                  SwitchListTile(
                    title: const Text("تحسين اكتشاف الأرقام العربية-الهندية", style: TextStyle(fontSize: 14)),
                    subtitle: const Text("سيتم التعرف على الأرقام بدون معالجة مسبقة للصورة", style: TextStyle(fontSize: 11, color: Colors.grey)),
                    secondary: const Icon(Icons.document_scanner, color: Colors.white70),
                    value: _isIndicEnhanceEnabled,
                    activeColor: Colors.green,
                    onChanged: (bool value) {
                      setDialogState(() => _isIndicEnhanceEnabled = value);
                      setState(() => _isIndicEnhanceEnabled = value);
                    },
                  ),
                  SwitchListTile(
                    title: const Text("تحسين لقراءة المكتوب بقلم أحمر", style: TextStyle(fontSize: 14)),
                    subtitle: const Text("سيتم مسح الصورة كما هي (أفضل للأقلام الزرقاء والسوداء)", style: TextStyle(fontSize: 11, color: Colors.grey)),
                    secondary: const Icon(Icons.palette, color: Colors.white70),
                    value: _isRedPenEnhanceEnabled,
                    activeColor: Colors.green,
                    onChanged: (bool value) {
                      setDialogState(() => _isRedPenEnhanceEnabled = value);
                      setState(() => _isRedPenEnhanceEnabled = value);
                    },
                  ),
                ],
              ),
              actions: [
                TextButton(
                  child: const Text("إغلاق", style: TextStyle(color: Colors.green, fontSize: 16)),
                  onPressed: () => Navigator.of(context).pop(),
                ),
              ],
            );
          },
        );
      },
    );
  }

  void _showSnackBar(String message) {
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(message)));
  }

  @override
  void dispose() {
    _gradeController.dispose();
    _cameraController.dispose();
    _textRecognizer.close();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    Color primaryPurple = const Color(0xFF7B1FA2);
    Color fieldColor = const Color(0xFF212121);

    return Scaffold(
      backgroundColor: const Color(0xFF4A148C),
      appBar: AppBar(
        title: const Text("برنامج إسقاط الدرجات بالأكواد"),
        centerTitle: true,
        backgroundColor: primaryPurple,
        leading: IconButton(
          icon: const Icon(Icons.settings, color: Colors.white),
          onPressed: _showSettingsDialog, 
        ),
        actions: [
          IconButton(
            icon: Icon(_isTorchOn ? Icons.flash_on : Icons.flash_off),
            onPressed: () {
              if (_isScanningActive) {
                _cameraController.toggleTorch();
                setState(() {
                  _isTorchOn = !_isTorchOn;
                });
              }
            }, 
          ),
        ],
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          children: [
            ElevatedButton(
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.green,
                minimumSize: const Size.fromHeight(50),
              ),
              onPressed: _isLoading ? null : _pickAndParseExcel, 
              child: _isLoading 
                ? const SizedBox(width: 24, height: 24, child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2))
                : const Text(
                    "اختر ملف الأكسيل الأصلي",
                    style: TextStyle(fontSize: 18, color: Colors.white),
                  ),
            ),
            const SizedBox(height: 8),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                color: fieldColor,
                borderRadius: BorderRadius.circular(8),
              ),
              child: Text(
                _fileName,
                style: const TextStyle(color: Colors.white),
                textAlign: TextAlign.center,
              ),
            ),
            const SizedBox(height: 12),

            const Align(
              alignment: Alignment.centerRight,
              child: Text(
                "اختر المادة :",
                style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
              ),
            ),
            const SizedBox(height: 4),
            Container(
              padding: const EdgeInsets.symmetric(horizontal: 12),
              decoration: BoxDecoration(
                color: fieldColor,
                borderRadius: BorderRadius.circular(8),
              ),
              child: DropdownButtonHideUnderline(
                child: DropdownButton<String>(
                  dropdownColor: fieldColor,
                  isExpanded: true,
                  hint: Text(
                    _subjects.isEmpty ? "يرجى اختيار ملف الأكسيل لجلب المواد" : "انقر لتحديد المادة المفتوحة ورصدها",
                    style: const TextStyle(color: Colors.grey),
                  ),
                  value: _selectedSubject,
                  items: _subjects.isEmpty ? null : _subjects
                      .map(
                        (sub) => DropdownMenuItem(
                          value: sub,
                          child: Text(sub, style: const TextStyle(color: Colors.white)),
                        ),
                      )
                      .toList(),
                  onChanged: (val) {
                    setState(() {
                      _selectedSubject = val;
                    });
                  },
                ),
              ),
            ),
            const SizedBox(height: 12),

            const Align(
              alignment: Alignment.centerRight,
              child: Text(
                "الرقم السري :",
                style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
              ),
            ),
            const SizedBox(height: 4),
            Container(
              width: double.infinity,
              padding: const EdgeInsets.all(12),
              decoration: BoxDecoration(
                color: fieldColor,
                borderRadius: BorderRadius.circular(8),
              ),
              child: Text(
                _secretIdResult,
                style: const TextStyle(color: Colors.white, fontSize: 18, fontWeight: FontWeight.bold),
                textAlign: TextAlign.center,
              ),
            ),
            const SizedBox(height: 12),

            Row(
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      const Text(
                        "الدرجة :",
                        style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
                      ),
                      const SizedBox(height: 4),
                      TextField(
                        controller: _gradeController,
                        keyboardType: TextInputType.number,
                        textAlign: TextAlign.center,
                        style: const TextStyle(color: Colors.white, fontSize: 18, fontWeight: FontWeight.bold),
                        decoration: InputDecoration(
                          fillColor: fieldColor,
                          filled: true,
                          border: OutlineInputBorder(
                            borderRadius: BorderRadius.circular(8),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 16),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.end,
                    children: [
                      const Text(
                        "العداد :",
                        style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold),
                      ),
                      const SizedBox(height: 4),
                      Container(
                        width: double.infinity,
                        padding: const EdgeInsets.all(14),
                        decoration: BoxDecoration(
                          color: fieldColor,
                          borderRadius: BorderRadius.circular(8),
                          // تم استخدام التوجيه الصريح هنا لحل مشكلة تعارض الاستدعاء لـ Border
                          border: painting.Border.all(color: Colors.grey.shade700, width: 1),
                        ),
                        child: Text(
                          "$_gradedStudents / $_totalStudents",
                          style: const TextStyle(
                            color: Colors.greenAccent,
                            fontWeight: FontWeight.bold,
                            fontSize: 16,
                          ),
                          textAlign: TextAlign.center,
                        ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
            const SizedBox(height: 20),

            ElevatedButton(
              style: ElevatedButton.styleFrom(
                backgroundColor: _isScanningActive ? Colors.red : Colors.green,
                minimumSize: const Size.fromHeight(50),
              ),
              onPressed: _toggleScanning, 
              child: Text(
                _isScanningActive ? "إيقاف المسح مؤقتاً" : "ابدأ المسح بالكاميرا",
                style: const TextStyle(fontSize: 18, color: Colors.white),
              ),
            ),
            const SizedBox(height: 8),
            ElevatedButton(
              style: ElevatedButton.styleFrom(
                backgroundColor: Colors.blueAccent,
                minimumSize: const Size.fromHeight(50),
              ),
              onPressed: (_selectedFilePath != null && _tempGradesCache.isNotEmpty) ? _saveAndExportToExcel : null, 
              child: const Text(
                "حفظ وتعديل الملف الأصلي",
                style: TextStyle(fontSize: 18, color: Colors.white),
              ),
            ),
            const SizedBox(height: 20),

            Container(
              width: double.infinity,
              height: 250,
              decoration: BoxDecoration(
                color: Colors.black,
                borderRadius: BorderRadius.circular(12),
                border: painting.Border.all(color: _isScanningActive ? Colors.greenAccent : Colors.grey, width: 2),
              ),
              child: ClipRRect(
                borderRadius: BorderRadius.circular(10),
                child: _isScanningActive
                    ? MobileScanner(
                        controller: _cameraController,
                        onDetect: _processCapturedImage,
                      )
                    : const Center(
                        child: Text(
                          "انقر فوق 'ابدأ المسح بالكاميرا' لتشغيل الفحص الحي",
                          style: TextStyle(color: Colors.white70),
                          textAlign: TextAlign.center,
                        ),
                      ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
