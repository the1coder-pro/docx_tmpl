import 'dart:async';
import 'dart:convert';

import 'package:archive/archive.dart';
import 'package:collection/collection.dart' show IterableExtension;
import 'package:flutter/foundation.dart' show kIsWeb;
import 'package:flutter/services.dart' show rootBundle;
import 'package:http/http.dart' as http;
import 'package:universal_io/io.dart'; // Replaces dart:io to prevent web compilation crashes
import 'package:xml/xml.dart';

import '../docx_tmpl.dart';
import 'docx_constants.dart';
import 'services/index.dart';

/// Main [DocxTmpl] class that supports Web, Mobile, and Desktop natively.
class DocxTmpl {
  /// .docx file path, can be local path, asset path, or http(s) url.
  final String? docxTemplate;

  /// Internal zip file object of the read .docx file.
  late Archive _zip;

  /// In-memory file registry mimicking the extracted directory structure.
  /// This completely removes the need for local temporary directories.
  final Map<String, List<int>> _virtualFiles = {};

  /// Holds parsed XML document parts.
  final Map<dynamic, XmlDocument> _parts = <dynamic, XmlDocument>{};

  /// Holds textual elements containing merge fields.
  final List<XmlElement> _instrTextChildren = [];

  /// Holds unique merge fields extracted from the document.
  List<String> mergedFields = [];

  int _imageCounter = 1;

  /// Static in-memory cache to support web offline document generation across instances
  static final Map<String, List<int>> _templateCache = {};

  DocxTmpl({required this.docxTemplate});

  /// Reads XML structures directly from the in-memory virtual filesystem.
  Future<List> __getTreeOfFile(XmlElement file) async {
    var type = file.getAttribute('PartName')!;
    // The zip archive stores paths without the leading slash
    var innerFile = type.replaceFirst('/', '');

    var fileBytes = _virtualFiles[innerFile];
    if (fileBytes == null) {
      throw Exception('Part $innerFile not found in the document archive.');
    }

    // Decode bytes to String and parse the XML
    String xmlStringData = utf8.decode(fileBytes);
    var parsedZi = XmlDocument.parse(xmlStringData);

    return [innerFile, parsedZi];
  }

  void _addContentTypeExtension(String ext) {
    if (_virtualFiles['[Content_Types].xml'] == null) return;
    var ctBytes = _virtualFiles['[Content_Types].xml']!;
    var ctStr = utf8.decode(ctBytes);
    if (!ctStr.contains('Extension="$ext"')) {
      String mime = ext == 'jpg' || ext == 'jpeg' ? 'image/jpeg' : 'image/$ext';
      String newTag = '<Default Extension="$ext" ContentType="$mime"/>';
      ctStr = ctStr.replaceFirst('</Types>', '$newTag</Types>');
      _virtualFiles['[Content_Types].xml'] = utf8.encode(ctStr);
    }
  }

  String _addRelationship(String partName, String targetPath) {
    var segments = partName.split('/');
    var fileName = segments.removeLast();
    segments.add('_rels');
    segments.add('$fileName.rels');
    String relsPath = segments.join('/');

    String relsStr;
    if (_virtualFiles.containsKey(relsPath)) {
      relsStr = utf8.decode(_virtualFiles[relsPath]!);
    } else {
      relsStr =
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }

    int maxId = 0;
    var matches = RegExp(r'Id="rId(\d+)"').allMatches(relsStr);
    for (var m in matches) {
      int id = int.parse(m.group(1)!);
      if (id > maxId) maxId = id;
    }
    String newRId = 'rId${maxId + 1}';

    String relativeTarget = targetPath;
    if (partName.startsWith('word/') && targetPath.startsWith('word/')) {
      relativeTarget = targetPath.substring(5);
    }

    String relTag =
        '<Relationship Id="$newRId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="$relativeTarget"/>';
    relsStr = relsStr.replaceFirst(
      '</Relationships>',
      '$relTag</Relationships>',
    );
    _virtualFiles[relsPath] = utf8.encode(relsStr);

    return newRId;
  }

  String _addExternalRelationship(String partName, String targetUrl) {
    var segments = partName.split('/');
    var fileName = segments.removeLast();
    segments.add('_rels');
    segments.add('$fileName.rels');
    String relsPath = segments.join('/');

    String relsStr;
    if (_virtualFiles.containsKey(relsPath)) {
      relsStr = utf8.decode(_virtualFiles[relsPath]!);
    } else {
      relsStr =
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    }

    int maxId = 0;
    var matches = RegExp(r'Id="rId(\d+)"').allMatches(relsStr);
    for (var m in matches) {
      int id = int.parse(m.group(1)!);
      if (id > maxId) maxId = id;
    }
    String newRId = 'rId${maxId + 1}';

    String relTag =
        '<Relationship Id="$newRId" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="$targetUrl" TargetMode="External"/>';
    relsStr = relsStr.replaceFirst(
      '</Relationships>',
      '$relTag</Relationships>',
    );
    _virtualFiles[relsPath] = utf8.encode(relsStr);

    return newRId;
  }

  String _generateDrawingXml(String rId, double widthPx, double heightPx) {
    int widthEmu = (widthPx * 9525).toInt();
    int heightEmu = (heightPx * 9525).toInt();
    int docPrId = DateTime.now().millisecondsSinceEpoch % 100000;
    return '''<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="$widthEmu" cy="$heightEmu"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="$docPrId" name="Picture $docPrId"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="$docPrId" name="Image"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="$rId" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="$widthEmu" cy="$heightEmu"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>''';
  }

  String _processImage(DocxImage img, String partName, String rawMatch) {
    String imageName = 'img_${_imageCounter++}.${img.extension}';
    String mediaPath = 'word/media/$imageName';
    _virtualFiles[mediaPath] = img.bytes;
    _addContentTypeExtension(img.extension);
    String relId = _addRelationship(partName, mediaPath);
    String drawingXml = _generateDrawingXml(relId, img.width, img.height);
    Iterable<Match> tags = RegExp(r'<[^>]+>').allMatches(rawMatch);
    return '</w:t>$drawingXml<w:t xml:space="preserve">' +
        tags.map((t) => t.group(0)!).join('');
  }

  String _processHyperlink(
    DocxHyperlink link,
    String partName,
    String rawMatch,
  ) {
    String relId = _addExternalRelationship(partName, link.url);
    String safeText = link.text
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;');
    String hyperlinkXml =
        '</w:t></w:r><w:hyperlink r:id="$relId"><w:r><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr><w:t>$safeText</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve">';
    Iterable<Match> tags = RegExp(r'<[^>]+>').allMatches(rawMatch);
    return hyperlinkXml + tags.map((t) => t.group(0)!).join('');
  }

  /// Retrieves a list of unique merge fields extracted from the .docx template.
  List<String> getMergeFields() {
    return mergedFields.toSet().toList()..remove('i');
  }

  /// Saves the modified document.
  ///
  /// On Native/Desktop: If [filepath] is provided, it saves to disk and returns the path (String).
  /// On Web (or if [filepath] is null): It returns the raw document bytes (List<int>).
  Future<dynamic> save([String? filepath]) async {
    final updatedArchive = Archive();

    // Iterate through virtual memory files and package them into the new Archive
    _virtualFiles.forEach((filename, content) {
      updatedArchive.addFile(ArchiveFile(filename, content.length, content));
    });

    // Encode the archive back into standard zip/docx format
    final encodedBytes = ZipEncoder().encode(updatedArchive);

    if (encodedBytes == null) {
      throw Exception('Failed to encode the document archive.');
    }

    if (kIsWeb || filepath == null) {
      // For Web, return raw bytes so the UI can trigger a browser download
      return encodedBytes;
    } else {
      // For Native OS, save to local storage
      final file = File(filepath);
      await file.create(recursive: true);
      await file.writeAsBytes(encodedBytes);
      return filepath;
    }
  }

  /// Writes the provided data to the merge fields in the document.
  Future<void> writeMergeFields({required Map<String, dynamic> data}) async {
    // Process all parsed parts (e.g. word/document.xml, headers, footers)
    for (var entry in _parts.entries) {
      String partName = entry.key;
      XmlDocument partXmlDoc = entry.value;

      var xml = partXmlDoc.toXmlString(pretty: false);

      final s = r'(?:<[^>]+>|\s)*';
      String kw(String word) => word.split('').join(s);

      // Helper: If a tag is the ONLY text in its Word paragraph, expand the bounds
      // to cleanly remove the entire paragraph and prevent blank lines/spaces.
      List<int> expandToParagraphIfEmpty(String xmlStr, int start, int end) {
        int pStart1 = xmlStr.lastIndexOf('<w:p>', start);
        int pStart2 = xmlStr.lastIndexOf('<w:p ', start);
        int pStart = pStart1 > pStart2 ? pStart1 : pStart2;

        int pEnd = xmlStr.indexOf('</w:p>', end);

        if (pStart != -1 && pEnd != -1) {
          int nextP1 = xmlStr.indexOf('<w:p>', pStart + 4);
          int nextP2 = xmlStr.indexOf('<w:p ', pStart + 4);
          int nextP = -1;
          if (nextP1 != -1 && nextP2 != -1) {
            nextP = nextP1 < nextP2 ? nextP1 : nextP2;
          } else {
            nextP = nextP1 != -1 ? nextP1 : nextP2;
          }

          if (nextP == -1 || nextP > start) {
            String pContent = xmlStr.substring(pStart, pEnd + 6);
            if (pContent.contains('<w:drawing') ||
                pContent.contains('<v:shape') ||
                pContent.contains('<w:pict')) {
              return [
                start,
                end,
              ]; // Keep the paragraph if it contains an image/shape
            }
            String pText = pContent
                .replaceAll(RegExp(r'<[^>]+>'), '')
                .replaceAll(RegExp(r'\s+'), '');
            String matchText = xmlStr
                .substring(start, end)
                .replaceAll(RegExp(r'<[^>]+>'), '')
                .replaceAll(RegExp(r'\s+'), '');
            if (pText == matchText) {
              return [pStart, pEnd + 6];
            }
          }
        }
        return [start, end];
      }

      // 1. Process for loops on the raw XML string
      String forKeywords = '(${kw("endfor")}|${kw("for")})';
      final forTagRegex = RegExp(
        r'\{' + s + r'%' + s + forKeywords + r'(.*?)' + r'%' + s + r'\}',
        caseSensitive: false,
        dotAll: true,
      );

      while (true) {
        var matches = forTagRegex.allMatches(xml).toList();
        if (matches.isEmpty) break;

        int innermostForIdx = -1;
        int endforIdx = -1;

        for (int i = 0; i < matches.length; i++) {
          String kwClean = matches[i]
              .group(1)!
              .replaceAll(RegExp(r'<[^>]+>|\s'), '')
              .toLowerCase();
          if (kwClean == 'for') {
            innermostForIdx = i;
            endforIdx = -1;
          } else if (kwClean == 'endfor' && innermostForIdx != -1) {
            endforIdx = i;
            break; // Found the innermost complete block
          }
        }

        if (innermostForIdx != -1 && endforIdx != -1) {
          var forMatch = matches[innermostForIdx];
          var endforMatch = matches[endforIdx];

          String conditionClean = forMatch
              .group(2)!
              .replaceAll(RegExp(r'<[^>]+>'), '')
              .trim();
          var parts = conditionClean.split(RegExp(r'\s+in\s+'));
          if (parts.length == 2) {
            String itemName = parts[0].trim();
            String listName = parts[1].trim();

            int trStartFor1 = xml.lastIndexOf('<w:tr>', forMatch.start);
            int trStartFor2 = xml.lastIndexOf('<w:tr ', forMatch.start);
            int trStartFor = trStartFor1 > trStartFor2
                ? trStartFor1
                : trStartFor2;
            int trEndFor = xml.indexOf('</w:tr>', forMatch.end);

            int prevTrEndFor = xml.lastIndexOf('</w:tr>', forMatch.start);
            bool isInsideRowFor = trStartFor != -1 && trStartFor > prevTrEndFor;

            int trStartEndfor1 = xml.lastIndexOf('<w:tr>', endforMatch.start);
            int trStartEndfor2 = xml.lastIndexOf('<w:tr ', endforMatch.start);
            int trStartEndfor = trStartEndfor1 > trStartEndfor2
                ? trStartEndfor1
                : trStartEndfor2;
            int trEndEndfor = xml.indexOf('</w:tr>', endforMatch.end);

            int prevTrEndEndfor = xml.lastIndexOf('</w:tr>', endforMatch.start);
            bool isInsideRowEndfor =
                trStartEndfor != -1 && trStartEndfor > prevTrEndEndfor;

            bool isSameTableRow =
                isInsideRowFor &&
                isInsideRowEndfor &&
                trStartFor == trStartEndfor &&
                trEndFor == trEndEndfor;

            var forBounds = expandToParagraphIfEmpty(
              xml,
              forMatch.start,
              forMatch.end,
            );
            var endforBounds = expandToParagraphIfEmpty(
              xml,
              endforMatch.start,
              endforMatch.end,
            );

            String loopContent;
            int replaceStart;
            int replaceEnd;

            if (isSameTableRow) {
              replaceStart = trStartFor;
              replaceEnd = trEndFor + 7; // length of </w:tr>

              int forB0 = forBounds[0] - replaceStart;
              int forB1 = forBounds[1] - replaceStart;
              int endforB0 = endforBounds[0] - replaceStart;
              int endforB1 = endforBounds[1] - replaceStart;

              String rowStr = xml.substring(replaceStart, replaceEnd);

              loopContent =
                  rowStr.substring(0, forB0) +
                  rowStr.substring(forB1, endforB0) +
                  rowStr.substring(endforB1);
            } else {
              replaceStart = forBounds[0];
              replaceEnd = endforBounds[1];
              loopContent = xml.substring(forBounds[1], endforBounds[0]);
            }

            var listData = data[listName];
            String duplicatedContent = "";

            if (listData is List) {
              final iterVarRegex = RegExp(
                r'\{' + s + r'\{' + r'(.*?)' + r'\}' + s + r'\}',
                dotAll: true,
              );

              final indexRegex = RegExp(
                r'\{' + s + r'i' + s + r'\}',
                dotAll: true,
              );

              // 1. Calculate the step (maximum index found in loop content)
              int step = 1;
              for (var m in iterVarRegex.allMatches(loopContent)) {
                String content = m
                    .group(1)!
                    .replaceAll(RegExp(r'<[^>]+>|\s'), '');
                var commaParts = content.split(',');
                if (commaParts.length == 2) {
                  int? sVal = int.tryParse(commaParts[1].trim());
                  if (sVal != null && sVal > step) step = sVal;
                }
              }

              for (int index = 0; index < listData.length; index += step) {
                String currentIterContent = loopContent
                    .replaceAllMapped(indexRegex, (iMatch) {
                      String rawMatch = iMatch.group(0)!;
                      // Calculate row/iteration number based on step
                      String strVal = ((index ~/ step) + 1).toString();
                      Iterable<Match> tags = RegExp(
                        r'<[^>]+>',
                      ).allMatches(rawMatch);
                      return strVal + tags.map((t) => t.group(0)!).join('');
                    })
                    .replaceAllMapped(iterVarRegex, (vMatch) {
                      String rawMatch = vMatch.group(0)!;
                      String cleanVarFull = vMatch
                          .group(1)!
                          .replaceAll(RegExp(r'<[^>]+>|\s'), '');

                      String cleanVar = cleanVarFull;
                      int offset = 1;
                      var commaParts = cleanVarFull.split(',');
                      if (commaParts.length == 2) {
                        cleanVar = commaParts[0].trim();
                        offset = int.tryParse(commaParts[1].trim()) ?? 1;
                      }

                      // Fetch item based on current iteration index + requested offset
                      int dataIndex = index + offset - 1;
                      var currentItemData = (dataIndex < listData.length)
                          ? listData[dataIndex]
                          : null;

                      String? strVal;
                      dynamic val;
                      if (currentItemData != null) {
                        if (cleanVar.startsWith('$itemName.')) {
                          String prop = cleanVar.substring(itemName.length + 1);
                          val = (currentItemData is Map)
                              ? currentItemData[prop]
                              : null;
                          strVal = val?.toString() ?? '';
                        } else if (cleanVar == itemName) {
                          val = currentItemData;
                          strVal = currentItemData.toString();
                        }
                      } else {
                        strVal = ''; // Put empty if data is out of bounds
                      }

                      if (val is DocxImage) {
                        return _processImage(val, partName, rawMatch);
                      } else if (val is DocxHyperlink) {
                        return _processHyperlink(val, partName, rawMatch);
                      } else if (strVal != null) {
                        strVal = strVal
                            .replaceAll('&', '&amp;')
                            .replaceAll('<', '&lt;')
                            .replaceAll('>', '&gt;');
                        Iterable<Match> tags = RegExp(
                          r'<[^>]+>',
                        ).allMatches(rawMatch);
                        return strVal + tags.map((t) => t.group(0)!).join('');
                      }
                      return rawMatch;
                    });
                duplicatedContent += currentIterContent;
              }
            }

            xml =
                xml.substring(0, replaceStart) +
                duplicatedContent +
                xml.substring(replaceEnd);
          } else {
            xml =
                xml.substring(0, forMatch.start) +
                xml.substring(forMatch.end, endforMatch.start) +
                xml.substring(endforMatch.end);
          }
        } else {
          break;
        }
      }

      // 2. Process conditional if/else blocks
      String ifKeywords = '(${kw("endif")}|${kw("else")}|${kw("if")})';
      final tagRegex = RegExp(
        r'\{' + s + r'%' + s + ifKeywords + r'(.*?)' + r'%' + s + r'\}',
        caseSensitive: false,
        dotAll: true,
      );

      while (true) {
        var matches = tagRegex.allMatches(xml).toList();
        if (matches.isEmpty) break;

        int innermostIfIdx = -1;
        int elseIdx = -1;
        int endifIdx = -1;

        for (int i = 0; i < matches.length; i++) {
          String kwClean = matches[i]
              .group(1)!
              .replaceAll(RegExp(r'<[^>]+>|\s'), '')
              .toLowerCase();
          if (kwClean == 'if') {
            innermostIfIdx = i;
            elseIdx = -1; // reset else when we see a new if
          } else if (kwClean == 'else' && innermostIfIdx != -1) {
            elseIdx = i;
          } else if (kwClean == 'endif' && innermostIfIdx != -1) {
            endifIdx = i;
            break; // Found the innermost complete block
          }
        }

        if (innermostIfIdx != -1 && endifIdx != -1) {
          var ifMatch = matches[innermostIfIdx];
          var elseMatch = elseIdx != -1 ? matches[elseIdx] : null;
          var endifMatch = matches[endifIdx];

          String conditionFieldRaw = ifMatch.group(2)!;
          String conditionClean = conditionFieldRaw
              .replaceAll(RegExp(r'<[^>]+>'), '')
              .trim();

          bool negate = false;
          if (conditionClean.startsWith('!')) {
            negate = true;
            conditionClean = conditionClean.substring(1).trim();
          }

          bool isTrue = false;
          final eqIdx = conditionClean.indexOf('==');
          if (eqIdx != -1) {
            String fieldRaw = conditionClean.substring(0, eqIdx).trim();
            String valRaw = conditionClean.substring(eqIdx + 2).trim();

            valRaw = valRaw
                .replaceAll('&quot;', '')
                .replaceAll('"', '')
                .replaceAll('”', '')
                .replaceAll('“', '')
                .replaceAll('&apos;', '')
                .replaceAll("'", '');

            var val = data[fieldRaw];
            isTrue = val?.toString() == valRaw;
          } else {
            String fieldRaw = conditionClean.trim();
            var val = data[fieldRaw];
            if (val != null) {
              if (val is bool) {
                isTrue = val;
              } else if (val is String) {
                isTrue =
                    val.isNotEmpty &&
                    val.toLowerCase() != 'false' &&
                    val != '0';
              } else if (val is num) {
                isTrue = val != 0;
              } else {
                isTrue = true;
              }
            }
          }

          if (negate) {
            isTrue = !isTrue;
          }

          var ifBounds = expandToParagraphIfEmpty(
            xml,
            ifMatch.start,
            ifMatch.end,
          );
          var elseBounds = elseMatch != null
              ? expandToParagraphIfEmpty(xml, elseMatch.start, elseMatch.end)
              : null;
          var endifBounds = expandToParagraphIfEmpty(
            xml,
            endifMatch.start,
            endifMatch.end,
          );

          String newContent;
          if (isTrue) {
            int trueStart = ifBounds[1];
            int trueEnd = elseBounds != null ? elseBounds[0] : endifBounds[0];
            newContent = xml.substring(trueStart, trueEnd);
          } else {
            if (elseBounds != null) {
              int falseStart = elseBounds[1];
              int falseEnd = endifBounds[0];
              newContent = xml.substring(falseStart, falseEnd);
            } else {
              newContent = "";
            }
          }

          // Modify the string sequentially for the current evaluated block
          xml =
              xml.substring(0, ifBounds[0]) +
              newContent +
              xml.substring(endifBounds[1]);
        } else {
          // No matching complete block found (e.g. trailing tags without 'endif')
          break;
        }
      }

      // 3. Process remaining global variable placeholders securely across XML layouts
      final globalVarRegex = RegExp(
        r'\{' + s + r'\{' + r'(.*?)' + r'\}' + s + r'\}',
        dotAll: true,
      );

      xml = xml.replaceAllMapped(globalVarRegex, (vMatch) {
        String rawMatch = vMatch.group(0)!;
        String rawVar = vMatch.group(1)!;
        String cleanVar = rawVar.replaceAll(RegExp(r'<[^>]+>|\s'), '');

        var val = data[cleanVar];

        if (val is DocxImage) {
          return _processImage(val, partName, rawMatch);
        } else if (val is DocxHyperlink) {
          return _processHyperlink(val, partName, rawMatch);
        }

        String strVal = val?.toString() ?? '';

        // HTML escape to prevent XML corruption
        strVal = strVal
            .replaceAll('&', '&amp;')
            .replaceAll('<', '&lt;')
            .replaceAll('>', '&gt;');

        Iterable<Match> tags = RegExp(r'<[^>]+>').allMatches(rawMatch);
        return strVal + tags.map((t) => t.group(0)!).join('');
      });

      // Save modified XML back to the virtual file system
      _virtualFiles[partName] = utf8.encode(xml);
    }
  }

  /// Parses the DOCX template file entirely in memory.
  Future<MergeResponse> parseDocxTmpl() async {
    try {
      List<int> bytes;

      // 1. Fetch File Bytes based on the source platform
      if (_templateCache.containsKey(docxTemplate)) {
        bytes = _templateCache[docxTemplate]!;
      } else if (isAssetFile) {
        List<int>? cachedBytes;
        File? cacheFile;
        if (!kIsWeb) {
          final tempDir = Directory.systemTemp;
          final fileName = 'docx_asset_cache_${docxTemplate.hashCode}.docx';
          cacheFile = File('${tempDir.path}/$fileName');
          if (await cacheFile.exists()) {
            cachedBytes = await cacheFile.readAsBytes();
          }
        }

        if (cachedBytes != null && cachedBytes.isNotEmpty) {
          bytes = cachedBytes;
        } else {
          final byteData = await rootBundle.load(docxTemplate!);
          bytes = byteData.buffer.asUint8List(
            byteData.offsetInBytes,
            byteData.lengthInBytes,
          );
          if (cacheFile != null) {
            await cacheFile.writeAsBytes(bytes);
          }
        }
        _templateCache[docxTemplate!] = bytes;
      } else if (isRemoteFile) {
        List<int>? cachedBytes;
        File? cacheFile;
        if (!kIsWeb) {
          final tempDir = Directory.systemTemp;
          final fileName = 'docx_cache_${docxTemplate.hashCode}.docx';
          cacheFile = File('${tempDir.path}/$fileName');
          if (await cacheFile.exists()) {
            cachedBytes = await cacheFile.readAsBytes();
          }
        }

        if (cachedBytes != null && cachedBytes.isNotEmpty) {
          bytes = cachedBytes;
        } else {
          final response = await http.get(Uri.parse(docxTemplate!));
          if (response.statusCode == 200) {
            bytes = response.bodyBytes;
            if (cacheFile != null) {
              await cacheFile.writeAsBytes(bytes);
            }
          } else {
            throw Exception(
              'Error downloading remote template. Status: ${response.statusCode}',
            );
          }
        }
        _templateCache[docxTemplate!] = bytes;
      } else {
        if (kIsWeb) {
          throw Exception(
            'Local file system paths are not supported on Flutter Web.',
          );
        }
        final ioFile = File(docxTemplate!);
        if (await ioFile.exists()) {
          bytes = await ioFile.readAsBytes();
        } else {
          throw Exception('Local file does not exist: $docxTemplate');
        }
        _templateCache[docxTemplate!] = bytes;
      }

      // 2. Decode the Zip Archive
      _zip = ZipDecoder().decodeBytes(bytes);

      // 3. Extract all files into the virtual in-memory system
      for (var zipInnerFile in _zip.files) {
        if (zipInnerFile.isFile) {
          _virtualFiles[zipInnerFile.name] = zipInnerFile.content as List<int>;
        }
      }

      // 4. Locate configuration map
      ArchiveFile? zippedFile = _zip.files.firstWhereOrNull(
        (zippedElement) => zippedElement.name == '[Content_Types].xml',
      );

      if (zippedFile == null) {
        throw Exception('Invalid document: [Content_Types].xml missing');
      }

      final contentBytes = _virtualFiles['[Content_Types].xml']!;
      final contentTypes = XmlDocument.parse(utf8.decode(contentBytes));

      // 5. Loop through xml document to identify target parts
      for (var file in contentTypes.findAllElements(
        'Override',
        namespace: "${NAMESPACES['ct']}",
      )) {
        var type = file.getAttribute(
          'ContentType',
          namespace: "${NAMESPACES['ct']}",
        );

        for (var contentTypePart in CONTENT_TYPES_PARTS) {
          if (type == contentTypePart) {
            var chunkResp = await __getTreeOfFile(file);
            _parts[chunkResp.first] = chunkResp.last;
          }
        }

        if (type == CONTENT_TYPE_SETTINGS) {
          var chunkResp = await __getTreeOfFile(file);
          _parts[chunkResp.first] = chunkResp.last;
        }

        // 6. Hunt for w:t text nodes to identify template tags
        for (var part in _parts.values) {
          for (var parent in part.findAllElements('w:t')) {
            _instrTextChildren.add(parent);
          }

          for (var instrChild in _instrTextChildren) {
            var chunkResult = templateParse(instrChild.innerText);
            mergedFields.addAll(chunkResult);
          }
        }
      }

      return MergeResponse(
        mergeStatus: MergeResponseStatus.Success,
        message: 'success',
      );
    } catch (e) {
      return MergeResponse(
        mergeStatus: MergeResponseStatus.Error,
        message: e.toString(),
      );
    }
  }

  DocxTemplateSource get templateSource {
    if (docxTemplate == null) {
      throw ArgumentError('docxTemplate cannot be null');
    }

    if (docxTemplate!.startsWith('http://') ||
        docxTemplate!.startsWith('https://')) {
      return DocxTemplateSource.remote;
    } else if (!docxTemplate!.startsWith('/') &&
        !docxTemplate!.contains(':\\')) {
      return DocxTemplateSource.asset;
    } else {
      return DocxTemplateSource.local;
    }
  }

  bool get isRemoteFile => templateSource == DocxTemplateSource.remote;
  bool get isAssetFile => templateSource == DocxTemplateSource.asset;
  bool get isLocalFile => templateSource == DocxTemplateSource.local;
}

/// Helper class to inject images into docx templates
class DocxImage {
  final List<int> bytes;
  final String extension;
  final double width;
  final double height;

  DocxImage({
    required this.bytes,
    this.extension = 'png',
    this.width = 100,
    this.height = 100,
  });
}

/// Helper class to inject hyperlinks into docx templates
class DocxHyperlink {
  final String url;
  final String text;

  DocxHyperlink({required this.url, required this.text});
}
