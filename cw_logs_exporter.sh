#!/bin/bash
# ==============================================================================
# CloudWatch Logs to Excel Exporter (for RHEL 9.8 / Bash)
#
# 機能概要:
#   1. スクリプト実行開始時に `aws login --remote` による認証状態をチェック。
#      未認証の場合は認証コマンドの案内を出力して終了。
#   2. CodeCommit用ロール等により各種AWS操作権限が不足している場合、
#      スイッチバックを促して終了、または '--auto-switchback' オプションにより
#      指定されたスイッチバック用シェルを自動で source して処理を続行。
#   3. CloudWatch Logs の全ログ（または指定プレフィックス/グループ）から、
#      指定時間範囲内のイベントを取得。
#   4. 各ロググループごとにシートを分け、ログストリームごとにブロック表示した
#      極めて美麗な Excel ファイル (.xlsx) を生成。
#      - Meiryo UI フォント統一
#      - セル幅の最適自動調整 & 文字折り返し
#      - 指定キーワード（ERROR, WARN等）のハイライト表示
#   5. RHEL 9.8 標準の Python 3 のみを使用し、追加インストール不要（外部ライブラリ依存ゼロ）。
# ==============================================================================

set -euo pipefail

# スクリプト自身のディレクトリ解決
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LIB_DIR="${SCRIPT_DIR}/lib"

# カラー出力設定（TTYの場合のみ）
if [ -t 1 ]; then
    RED='\033[0;31m'
    GREEN='\033[0;32m'
    YELLOW='\033[1;33m'
    BLUE='\033[0;34m'
    CYAN='\033[0;36m'
    BOLD='\033[1m'
    NC='\033[0m' # No Color
else
    RED=''
    GREEN=''
    YELLOW=''
    BLUE=''
    CYAN=''
    BOLD=''
    NC=''
fi

log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1" >&2
}

print_header() {
    echo -e "${CYAN}================================================================================${NC}"
    echo -e "${BOLD}${CYAN}   CloudWatch Logs Excel Exporter (RHEL 9.8 Edition)${NC}"
    echo -e "${CYAN}================================================================================${NC}"
}

usage() {
    cat << EOF
使用方法:
  $(basename "$0") [オプション]

必須/推奨オプション:
  -b, --start-time <TIME>       ログ取得開始日時 (デフォルト: 終了日時の1時間前)
                                例: "2026-10-04 00:00:00", "1h", "24h", "7d"
  -e, --end-time <TIME>         ログ取得終了日時 (デフォルト: 現在時刻)
                                例: "2026-10-04 06:00:00"
  -H, --highlight <WORDS>       強調（ハイライト）表示する文言 (カンマ区切り)
                                例: "ERROR,WARN,FATAL,Exception"

Excel出力オプション:
  -o, --output <PATH>           出力先 Excel ファイルパス (.xlsx)
                                デフォルト: ./cloudwatch_logs_YYYYMMDD_HHMMSS.xlsx

ロググループ絞り込みオプション:
  -p, --log-group-prefix <PRE>  対象ロググループ名のプレフィックス (例: /aws/lambda/)
  -g, --log-groups <NAMES>      対象ロググループ名をカンマ区切りで個別指定

認証 & スイッチバック設定:
  -a, --auto-switchback         AWS操作権限が不足している場合に、指定された
                                スイッチバック用シェルを自動で source して続行する。
                                (未指定時は警告メッセージを表示して終了)
  -s, --switchback-script <PATH>
                                スイッチバック用シェルのファイル配置場所
                                デフォルト: 環境変数 AWS_SWITCHBACK_SCRIPT または ./switchback.sh

AWS CLI 設定:
  -r, --region <REGION>         AWS リージョン (例: ap-northeast-1)
      --profile <PROFILE>       使用する AWS CLI プロファイル名

テスト用:
      --mock-data <JSON_PATH>   AWS API の代わりにモック JSON データを使用

その他:
  -h, --help                    このヘルプメッセージを表示して終了

EOF
    exit 0
}

# デフォルト設定値
START_TIME=""
END_TIME=""
OUTPUT_FILE=""
HIGHLIGHT_WORDS=""
LOG_GROUP_PREFIX=""
SPECIFIC_LOG_GROUPS=""
AUTO_SWITCHBACK=false
SWITCHBACK_SCRIPT="${AWS_SWITCHBACK_SCRIPT:-${SCRIPT_DIR}/switchback.sh}"
REGION=""
PROFILE=""
MOCK_DATA=""

# 引数解析 (getopt / getopts / 手動パース)
while [[ $# -gt 0 ]]; do
    case "$1" in
        -b|--start-time)
            START_TIME="$2"
            shift 2
            ;;
        -e|--end-time)
            END_TIME="$2"
            shift 2
            ;;
        -o|--output)
            OUTPUT_FILE="$2"
            shift 2
            ;;
        -H|--highlight)
            HIGHLIGHT_WORDS="$2"
            shift 2
            ;;
        -p|--log-group-prefix)
            LOG_GROUP_PREFIX="$2"
            shift 2
            ;;
        -g|--log-groups)
            SPECIFIC_LOG_GROUPS="$2"
            shift 2
            ;;
        -a|--auto-switchback)
            AUTO_SWITCHBACK=true
            shift 1
            ;;
        -s|--switchback-script)
            SWITCHBACK_SCRIPT="$2"
            shift 2
            ;;
        -r|--region)
            REGION="$2"
            shift 2
            ;;
        --profile)
            PROFILE="$2"
            shift 2
            ;;
        --mock-data)
            MOCK_DATA="$2"
            shift 2
            ;;
        -h|--help)
            usage
            ;;
        *)
            log_error "未知のオプションです: $1"
            echo "詳細は '$(basename "$0") --help' を実行してください。"
            exit 1
            ;;
    esac
done

print_header

# ------------------------------------------------------------------------------
# 1. 動作環境前提チェック (AWS CLI, Python 3)
# ------------------------------------------------------------------------------
if [ -z "${MOCK_DATA}" ]; then
    if ! command -v aws >/dev/null 2>&1; then
        log_error "AWS CLI ('aws' コマンド) がインストールされていないか、PATHが通っていません。"
        exit 1
    fi
fi

if ! command -v python3 >/dev/null 2>&1; then
    log_error "Python 3 ('python3' コマンド) がインストールされていないか、PATHが通っていません。"
    exit 1
fi

# ------------------------------------------------------------------------------
# 2. 認証チェック (`aws login --remote` が実行されているか)
# ------------------------------------------------------------------------------
if [ -z "${MOCK_DATA}" ]; then
    log_info "AWS 認証状態を確認しています (aws sts get-caller-identity) ..."

    AWS_CMD_BASE=("aws")
    if [ -n "${REGION}" ]; then
        AWS_CMD_BASE+=("--region" "${REGION}")
    fi
    if [ -n "${PROFILE}" ]; then
        AWS_CMD_BASE+=("--profile" "${PROFILE}")
    fi

    # STS 呼び出しによる認証判定
    set +e
    STS_OUTPUT=$("${AWS_CMD_BASE[@]}" sts get-caller-identity --output json 2>&1)
    STS_EXIT_CODE=$?
    set -e

    if [ ${STS_EXIT_CODE} -ne 0 ]; then
        echo ""
        echo -e "${RED}================================================================================${NC}"
        log_error "AWS にログインされていません (認証トークンが無効または未設定です)。"
        echo ""
        echo -e "エラー詳細: ${STS_OUTPUT}"
        echo ""
        echo -e "${BOLD}${YELLOW}事前に以下のコマンドを実行して認証を完了させてから、再実行してください:${NC}"
        echo ""
        echo -e "    ${BOLD}${CYAN}aws login --remote${NC}"
        echo ""
        echo -e "${RED}================================================================================${NC}"
        exit 1
    fi

    CURRENT_ARN=$(echo "${STS_OUTPUT}" | grep -o '"Arn": *"[^"]*"' | cut -d'"' -f4 || echo "Unknown")
    log_info "AWS 認証を確認しました。現在の Identity ARN: ${CYAN}${CURRENT_ARN}${NC}"

    # --------------------------------------------------------------------------
    # 3. 操作権限チェック & スイッチバック処理
    # --------------------------------------------------------------------------
    # CodeCommit用ロール等で各種AWS操作権限（CloudWatch Logsアクセス）が不足していないか確認
    log_info "各種 AWS (CloudWatch Logs) への操作権限を確認しています..."

    check_aws_permissions() {
        set +e
        PERM_CHECK_OUTPUT=$("${AWS_CMD_BASE[@]}" logs describe-log-groups --max-items 1 --output json 2>&1)
        PERM_CHECK_EXIT=$?
        set -e
        return ${PERM_CHECK_EXIT}
    }

    if ! check_aws_permissions; then
        echo ""
        echo -e "${YELLOW}================================================================================${NC}"
        log_warn "現在の操作権限では AWS (CloudWatch Logs) への操作が不可能です (AccessDenied)。"
        echo -e "エラー内容: ${PERM_CHECK_OUTPUT}"
        echo ""
        echo -e "※ CodeCommit 専用ロール等の限定権限が適用されている可能性があります。"
        echo -e "${YELLOW}================================================================================${NC}"

        if [ "${AUTO_SWITCHBACK}" = true ]; then
            log_info "'--auto-switchback' が指定されているため、自動的にスイッチバックを実行します。"
            
            if [ ! -f "${SWITCHBACK_SCRIPT}" ]; then
                log_error "スイッチバック用シェルが存在しません: ${SWITCHBACK_SCRIPT}"
                echo "正しいファイル配置場所を '--switchback-script <PATH>' で指定してください。"
                exit 3
            fi

            log_info "スイッチバック用シェルを source します: ${SWITCHBACK_SCRIPT}"
            # shellcheck source=/dev/null
            source "${SWITCHBACK_SCRIPT}"
            log_info "スイッチバック用シェルの読み込みが完了しました。"

            # スイッチバック後の権限再チェック
            log_info "スイッチバック後の操作権限を再検証しています..."
            if ! check_aws_permissions; then
                echo ""
                log_error "スイッチバックを実行しましたが、依然として AWS 操作権限が不足しています。"
                echo -e "エラー内容: ${PERM_CHECK_OUTPUT}"
                echo "スイッチバック用シェルの設定内容や切り替え先ロールの権限を確認してください。"
                exit 4
            fi
            log_info "${GREEN}スイッチバックに成功しました！AWS操作権限を確認しました。${NC}"
        else
            echo ""
            echo -e "${BOLD}${RED}[操作権限エラー]${NC}"
            echo -e "スイッチバック用シェルを実行して、AWS操作用ロールへ切り替えてから再実行してください。"
            echo ""
            echo -e "  スイッチバック用シェル配置場所: ${CYAN}${SWITCHBACK_SCRIPT}${NC}"
            echo -e "  手動実行例: ${BOLD}${CYAN}source \"${SWITCHBACK_SCRIPT}\"${NC}"
            echo ""
            echo -e "${BOLD}${YELLOW}※ 本スクリプト実行時に '--auto-switchback' (-a) オプションを指定すると、${NC}"
            echo -e "${BOLD}${YELLOW}   自動的に上記シェルを source してスイッチバックを実行し処理を続行できます。${NC}"
            echo ""
            echo -e "${RED}================================================================================${NC}"
            exit 2
        fi
    else
        log_info "${GREEN}AWS 操作権限を確認しました (CloudWatch Logs へのアクセス可能)。${NC}"
    fi
else
    log_warn "モックモード (--mock-data) で実行中のため、AWS認証および権限チェックをスキップします。"
fi

# ------------------------------------------------------------------------------
# 4. CloudWatch Logs ログ取得 & Excel レポート生成
# ------------------------------------------------------------------------------
log_info "CloudWatch Logs 取得および Excel 生成処理を開始します..."

PYTHON_ARGS=("${LIB_DIR}/cw_exporter.py")

if [ -n "${START_TIME}" ]; then
    PYTHON_ARGS+=("-b" "${START_TIME}")
fi
if [ -n "${END_TIME}" ]; then
    PYTHON_ARGS+=("-e" "${END_TIME}")
fi
if [ -n "${OUTPUT_FILE}" ]; then
    PYTHON_ARGS+=("-o" "${OUTPUT_FILE}")
fi
if [ -n "${HIGHLIGHT_WORDS}" ]; then
    PYTHON_ARGS+=("-H" "${HIGHLIGHT_WORDS}")
fi
if [ -n "${LOG_GROUP_PREFIX}" ]; then
    PYTHON_ARGS+=("-p" "${LOG_GROUP_PREFIX}")
fi
if [ -n "${SPECIFIC_LOG_GROUPS}" ]; then
    PYTHON_ARGS+=("-g" "${SPECIFIC_LOG_GROUPS}")
fi
if [ -n "${REGION}" ]; then
    PYTHON_ARGS+=("-r" "${REGION}")
fi
if [ -n "${PROFILE}" ]; then
    PYTHON_ARGS+=("--profile" "${PROFILE}")
fi
if [ -n "${MOCK_DATA}" ]; then
    PYTHON_ARGS+=("--mock-data" "${MOCK_DATA}")
fi

# Python スクリプト実行 (親シェルの環境変数は自動継承される)
python3 "${PYTHON_ARGS[@]}"

EXIT_CODE=$?
if [ ${EXIT_CODE} -eq 0 ]; then
    echo ""
    echo -e "${CYAN}================================================================================${NC}"
    log_info "${GREEN}${BOLD}すべての処理が正常に完了しました。${NC}"
    echo -e "${CYAN}================================================================================${NC}"
else
    log_error "Excel 生成処理中にエラーが発生しました (終了コード: ${EXIT_CODE})。"
    exit ${EXIT_CODE}
fi
