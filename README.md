# VBA
VBAマクロ
package jp.sumitclub.batch.mail;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.mail.Flags.Flag;
import javax.mail.Header;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;

public class UndeliveredMailReceiver extends BatchAction<Message> {

    /** ロガー */
    private static final Logger LOGGER = LoggerManager.get(UndeliveredMailReceiver.class);

    /** メール送受信用設定オブジェクト */
    private static final MailConfigEx MAIL_CONFIG = SystemRepository.get("mailConfig");

    /** メールヘッダからMessage-IDを検索する場合の完全一致正規表現（フォーマット：<メール送信要求ID@ドメイン>） */
    private static final Pattern HEADER_PATTERN = Pattern.compile("^<(.+)@" + MAIL_CONFIG.getMailServerDomain() + ">$");

    /** メール本文からMessage-IDを検索する場合の部分一致正規表現（フォーマット：<メール送信要求ID@ドメイン>） */
    private static final Pattern BODY_PATTERN = Pattern.compile("<(.+)@" + MAIL_CONFIG.getMailServerDomain() + ">");

    /** {@inheritDoc} */
    @Override
    public Result handle(Message message, ExecutionContext ctx) {

        try {
            // メールのMessage-IDを解析し、該当メールに対する配信履歴のメール送信要求IDの候補を抽出する。
            List<String> mailRequestIDs = parseMailRequestIDs(message);

            // 有効なメール送信要求IDとして配信履歴を更新できるまで全ての候補を処理する
            MailRequestTableEx mailRequestTable = SystemRepository.get("mailRequestTable");
            String statusUndelivered = MAIL_CONFIG.getStatusUndelivered();
            for (String mailRequestId : mailRequestIDs) {
                if (mailRequestTable.updateUndeliveredStatus(mailRequestId, statusUndelivered) > 0) {
                    break;
                }
            }

            // メールを処理済みとして削除フラグを設定する。
            message.setFlag(Flag.DELETED, true);

        } catch (MessagingException me) {

            // 受信エラー時のバッチ終了フラグをチェック
            if (MAIL_CONFIG.getReceiveErrorStoppingFlg()) {
                throw new ProcessAbnormalEnd(
                        MAIL_CONFIG.getReceiveAbnormalEndExitCode(), me,
                        MAIL_CONFIG.getReceiveFailureCode());
            }

            // エラー時でも終了させない場合は、WARNログを出力
            LOGGER.logWarn(MessageUtil.createMessage(
                    MessageLevel.WARN, MAIL_CONFIG.getReceiveFailureCode()).formatMessage());
        }
        return new Result.Success();
    }

    /**
     * メールメッセージを解析して該当する配信履歴のメール送信要求IDの候補を抽出する。
     *  - 全メールヘッダ
     *  - メール本文
     *  - 添付ファイル
     * @param message メールメッセージ
     * @return メール送信要求ID候補リスト
     */
    @SuppressWarnings("unchecked")
    protected List<String> parseMailRequestIDs(Message message) throws MessagingException {
        String domain = MAIL_CONFIG.getMailServerDomain();
        List<String> resultList = new ArrayList<>();

        // 全メールヘッダを検索
        Enumeration<Header> allHeaders = message.getAllHeaders();
        while (allHeaders.hasMoreElements()) {
            Matcher headerMatcher = HEADER_PATTERN.matcher(allHeaders.nextElement().getValue());
            if (headerMatcher.matches()) {
                resultList.add(headerMatcher.group(1));
            }
        }

        // メール本文を検索
        try {
            Object content = message.getContent();
            
            // 添付ファイルあり
            if (content instanceof Multipart) {
                final Multipart multiPart = (Multipart) content;
                for (int i = 0; i < multiPart.getCount(); i++) {
                    final Part part = multiPart.getBodyPart(i);
                    final String disposition = part.getDisposition();
                    if (Part.ATTACHMENT.equals(disposition) || Part.INLINE.equals(disposition)) {
                        
                        try (BufferedReader br = new BufferedReader(new InputStreamReader(part.getInputStream()))) {
                            br.lines().filter(this::)
                        }
                        
                    } else {
                        System.out.println("メール本文(添付ファイル付) ["
                                + part.getContent().toString()
                                + "]");
                    }
                }
                
                
                Matcher contentMatcher = BODY_PATTERN.matcher(content.toString());
                while (contentMatcher.find()){
                    resultList.add(contentMatcher.group(1));
                }
                
                
            } else {
                // 添付ファイルなし
                try {
                    Matcher contentMatcher = BODY_PATTERN.matcher(message.getContent().toString());
                    while (contentMatcher.find()){
                        resultList.add(contentMatcher.group(1));
                    }
                } catch (MessagingException | IOException e) {
                    LOGGER.logWarn("メール本文取得失敗", e);
                }
            }
        } catch (IOException ioe) {
            LOGGER.logWarn("メール本文の取得に失敗しました。", ioe);
        }

        return resultList;
    }

    /** {@inheritDoc} */
    @Override
    public DataReader<Message> createReader(ExecutionContext ctx) {
        MailMessageDataReader dataReader = new MailMessageDataReader();
        dataReader.setMailPOP3Config(SystemRepository.get("mailPOP3Config"));
        return dataReader;
    }
}
