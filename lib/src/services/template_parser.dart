/// parse the simple jinja2-like template from element tag text
/// {{...}}
List<String> templateParse(String text) {
  List<String> fields = [];

  // Extract {{ field }} or {{ field.subfield }} ignoring spaces
  final RegExp re = RegExp(
    r'\{\{\s*([\w\.]+)\s*\}\}',
    caseSensitive: true,
    multiLine: true,
  );
  Iterable<Match> matches = re.allMatches(text);
  if (matches.isNotEmpty) {
    for (var match in matches) {
      fields.add(match.group(1)!);
    }
  }

  // Extract conditions from if blocks
  final RegExp ifRe = RegExp(
    r'\{%\s*if\s+!?\s*(\w+)',
    caseSensitive: false,
    multiLine: true,
  );
  Iterable<Match> ifMatches = ifRe.allMatches(text);
  for (var match in ifMatches) {
    fields.add(match.group(1)!);
  }

  // Extract loop list names from for blocks
  final RegExp forRe = RegExp(
    r'\{%\s*for\s+(\w+)\s+in\s+([\w\.]+)\s*%\}',
    caseSensitive: false,
    multiLine: true,
  );
  Iterable<Match> forMatches = forRe.allMatches(text);
  for (var match in forMatches) {
    fields.add(match.group(2)!);
  }

  return fields;
}
