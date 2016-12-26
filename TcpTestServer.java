import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * TCPテストサーバ。
  * @since 1.0
 */
public class TcpTestServer extends Thread implements Closeable {

    /** サーバ文字コード */
    public static final Charset CHARSET = Charset.forName("US-ASCII");

    /** サーバ停止コマンド */
    public static final String STOP_SERVER = "STOP_SERVER";

    /** デフォルトポート番号 */
    public static final int DEFAULT_PORT_NO = 15210;

    /** サーバソケット */
    private ServerSocket serverSocket;

    /** 稼働中のサーバポート番号 */
    private int portNumber;

    /** サーバ稼働フラグ */
    private boolean continuous = false;

    /** 子スレッドリスト */
    private List<TcpTestServerThread> childThreads = new ArrayList<>();

    /** アクティブスレッド数 */
    private AtomicInteger activeThreadCount = new AtomicInteger(0);

    /** WAIT時間 */
    private int waitTime = 0;

    /** クローズモード（ステートレス通信にするためのフラグ） */
    private boolean closeMode = false;

    /**
     * コンストラクタ。
     */
    public TcpTestServer() {
        this(DEFAULT_PORT_NO);
    }

    /**
     * コンストラクタ。
     * @param portNumber サーバポート番号
     */
    public TcpTestServer(int portNumber) {
        setName("TCP Server Thread");
        this.portNumber = portNumber;
    }

    /**
     * メインメソッド。
     * @param args args
     * @throws Exception Exception
     */
    public static void main(String[] args) throws Exception {
        int portNumber = DEFAULT_PORT_NO;

        if (args.length > 0) {
            portNumber = Integer.parseInt(args[0]);
        }
        try (TcpTestServer server = new TcpTestServer(portNumber)) {
            String command = System.getProperty("command");
            if (STOP_SERVER.equals(command)) {
                server.sendStopCommand(true);
            } else {
                server.setCloseMode(Boolean.getBoolean("closeMode"));
                server.prepareServer();
            }
        }
    }

    /**
     * サーバ稼働フラグを取得する。
     * @return サーバ稼働フラグ
     */
    public boolean isContinuous() {
        return continuous;
    }

    /** {@inheritDoc} */
    @Override
    public void run() {
        prepareServer();
    }

    /**
     * サーバスタート。
     * @throws IOException IOException
     */
    public void startServer() {
        // サーバスレッド開始
        start();

        // サーバが起動完了するまで待機する
        while (!continuous) {
            wait(1);
        }
    }

    /**
     * サーバ準備。
     * @throws IOException IOException
     */
    public void prepareServer() {
        try {
            try (ServerSocket serverSocket = new ServerSocket(portNumber)) {
                this.serverSocket = serverSocket;
                serverSocket.setReuseAddress(true);
                System.out.println("TcpTestServer起動[port:" + serverSocket.getLocalPort() + "]");
                continuous = true;
                while (true) {
                    TcpTestServerThread childThread = new TcpTestServerThread(serverSocket.accept());
                    if (!continuous) { // リクエスト受信時点でサーバ稼働フラグがfalseであれば停止する
                        break;
                    }
                    childThreads.add(childThread);
                    childThread.start();
                }
                System.out.println("TcpTestServer停止[port:" + serverSocket.getLocalPort() + "]");
            }
        } catch (IOException ioe) {
            throw new RuntimeException(ioe);
        }
    }

    /**
     * サーバ停止。
     * @return サーバ停止成否
     */
    public boolean stopServer() {
        if (continuous) {
            continuous = false;
            System.out.println(STOP_SERVER + " command received.");
            if (!sendStopCommand(false)) {
                return false;
            }
        }

        // サーバが停止完了するまで待機
        while (serverSocket != null && !serverSocket.isClosed()) {
            wait(1);
        }
        return true;
    }

    /**
     * サーバ停止コマンドを送信する。
     * @param  isDefinite 停止電文の送付有無（true:送信あり, false:送信なし(停止用コネクションオープンのみ)）
     * @return 送信成否（true:送信成功, false:送信失敗）
     */
    private boolean sendStopCommand(boolean isDefinite) {
        try (Socket socket = new Socket("localhost", portNumber)) { // ServerSocket#acceptでの待機処理を外す
            if (isDefinite) {
                socket.getOutputStream().write(STOP_SERVER.getBytes(CHARSET));
            }
        } catch (IOException e) {
            System.err.println("TCPテストサーバの停止に失敗");
            e.printStackTrace();
            return false;
        }
        return true;
    }

    /**
     * サーバ停止。
     * @return サーバ停止成否
     */
    public int getAliveThreadCount() {
        int count = 0;
        synchronized (this) {
            for (TcpTestServerThread thread : childThreads) {
                count = thread.isAlive() ? count + 1 : count;
            }
        }
        return count;
    }

    /**
     * 処理を待機する。
     * @param time 待機時間
     */
    private static void wait(int time) {
        try {
            sleep(time);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * クライアントとの1コネクション単位の処理を担うスレッドクラス
     * @since 1.0
     */
    private class TcpTestServerThread extends Thread {

        /** クライアントとのコネクション */
        private Socket socket = null;

        /**
         * コンストラクタ。
         * @param socket socket
         */
        public TcpTestServerThread(Socket socket) {
            this.socket = socket;
            setName("TCP Server Child Thread[" + socket.getInetAddress().getHostAddress()
                    + ":" + socket.getPort() + "]");
        }

        /** {@inheritDoc} */
        @Override
        public void run() {
            System.out.println("new connection thread started. " + getName() + " Thread-State(active/total)"
                    + ":(" + activeThreadCount.get() + "/" + childThreads.size() + ")");

            try (Socket socket = this.socket;
                    InputStream inputStream = new BufferedInputStream(socket.getInputStream());
                    OutputStream outputStream = socket.getOutputStream()) {

                while (true) {
                    String line = getRequestString(inputStream);

                    // 注意：このログ出力は、UTのログ確認ケースに使用するため変更しないこと
                    System.out.println(createServerLog(socket.getInetAddress().getHostAddress(), socket.getPort(), line)
                            + " request received. Thread-State(active/total)"
                            + ":(" + activeThreadCount.incrementAndGet() + "/" + childThreads.size() + ")");

                    // リクエスト受信からレスポンス送信までWAIT設定がされていたら待機
                    if (waitTime > 0) {
                        TcpTestServer.wait(waitTime);
                    }

                    // パフォーマンス優先設定（空レスポンス設定）の場合を除き、リクエストに応じたレスポンスを送信
                    byte[] responseLine = new byte[0];
                    if (closeMode) {
                        outputStream.close();
                    } else {
                        responseLine = createResponse(line);
                        outputStream.write(responseLine);
                    }
                    System.out.println(
                            createServerLog(socket.getInetAddress().getHostAddress(), socket.getPort(),
                                    new String(responseLine))
                                    + " response sended. Thread-State(active/total)"
                                    + ":(" + activeThreadCount.decrementAndGet() + "/" + childThreads.size() + ")");

                    // リクエストデータがサーバ停止コマンドだったら停止する
                    if (STOP_SERVER.equals(line)) {
                        stopServer();
                        break;
                    }

                    // サーバ停止フラグ or クライアントとのSocketが閉じていたらスレッド終了
                    // クライアントのcloseはサーバのSocekt#isClosedに反映されないため、空電文で判定する
                    if (closeMode || !continuous || line == null || line.isEmpty()) {
                        break;
                    }
                }
            } catch (IOException e) {
                System.err.println("通信時にTCPテストサーバで障害発生");
                e.printStackTrace();
                throw new RuntimeException("通信時にTCPテストサーバで障害発生", e);
            } finally {
                synchronized (childThreads) {
                    childThreads.remove(this);
                }
            }
        }

        /**
         * TCPレスポンス受信.
         *
         * @param inputStream inputStream
         * @return レスポンスメッセージ
         */
        private String getRequestString(InputStream inputStream) {
            int tmpLength = 0;
            int available = 0;
            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            try {
                byte[] tmpBytes = new byte[socket.getReceiveBufferSize()];
                do {
                    tmpLength = inputStream.read(tmpBytes);
                    if (tmpLength == -1) {
                        break;
                    }
                    baos.write(tmpBytes, 0, tmpLength);
                    available = inputStream.available();
                } while (available > 0);

            } catch (IOException ioe) {
                if (!ioe.getMessage().toLowerCase().contains("connection reset")) {
                    System.err.println(getName() + " failed to recieve TCP response.");
                    ioe.printStackTrace();
                }
            }
            return new String(baos.toByteArray(), CHARSET);
        }
    }

    /**
     * サーバログメッセージを生成する。
     * @param address アドレス
     * @param port ポート番号
     * @param msg メッセージ
     * @return サーバログメッセージ
     */
    public static String createServerLog(String address, int port, String msg) {
        return "(" + address + ":" + port + ")[" + msg + "]";
    }

    /**
     * 要求電文に応じた応答電文を作成する。
     * @param request 要求電文
     * @return 応答電文
     */
    public static byte[] createResponse(String request) {
        return request.getBytes(CHARSET);
    }

    /**
     * サーバがリクエストに応答するまでのWAIT時間（デフォルト:0秒）を設定する。
     * @param waitTime wait時間
     */
    public void setWaitTime(int waitTime) {
        this.waitTime = waitTime;
    }

    /**
     * ステートレス通信にするためのフラグを設定する。
     * trueを設定すると、サーバが応答する都度、Socketをクローズする。
     * （デフォルト：false[クローズしない]）
     * @param closeMode closeMode
     */
    public void setCloseMode(boolean closeMode) {
        this.closeMode = closeMode;
    }

    /** {@inheritDoc} */
    @Override
    public void close() throws IOException {
        stopServer();
    }
}
