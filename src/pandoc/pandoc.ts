import {spawn} from "child_process";
import {PandocJson} from "src/pandoc/pandoc-json";

export function pandoc(src, args): Promise<string> {
    return new Promise((resolve, reject) => {
        let stdout = ""
        let stderr = ""

        let pandocProcess = spawn('pandoc', args);

        pandocProcess.stdin.end(src, 'utf-8');

        pandocProcess.stdout.on('data', (data) => {
            stdout += data
        });

        pandocProcess.stderr.on('data', (data) => {
            stderr += data
        });

        pandocProcess.on('exit', function (code) {
            if (stderr.length) {
                console.error("There was some pandoc warnings along the way:")
                console.error(stderr)
            }

            if (code == 0) {
                resolve(stdout)
            } else {
                reject(new Error("Pandoc returned non-zero exit code"))
            }
        });
    })
}

export async function markdownToPandocJson(markdown: string, flags: string[]) {
    let meta = await pandoc(markdown, ["-f", "markdown", "-t", "json", ...flags])
    return JSON.parse(meta) as PandocJson
}

export async function pandocJsonToDocx(pandocJson: PandocJson, flags: string[]) {
    return await pandoc(JSON.stringify(pandocJson), ["-f", "json", "-t", "docx", ...flags])
}

export const tokenClasses = [
    "KeywordTok",
    "NormalTok",
    "OperatorTok",
    "DataTypeTok",
    "PreprocessorTok",
    "DecValTok",
    "BaseNTok",
    "FloatTok",
    "ConstantTok",
    "CharTok",
    "SpecialCharTok",
    "StringTok",
    "VerbatimStringTok",
    "SpecialStringTok",
    "ImportTok",
    "CommentTok",
    "DocumentationTok",
    "AnnotationTok",
    "CommentVarTok",
    "OtherTok",
    "FunctionTok",
    "VariableTok",
    "ControlFlowTok",
    "BuiltInTok",
    "ExtensionTok",
    "AttributeTok",
    "RegionMarkerTok",
    "InformationTok",
    "WarningTok",
    "AlertTok",
    "ErrorTok"
]