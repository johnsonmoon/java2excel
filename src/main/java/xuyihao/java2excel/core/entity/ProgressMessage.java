package xuyihao.java2excel.core.entity;

/**
 * 操作进度
 * <p>
 * Created by Xuyh at 2016/07/26 上午 11:15.
 */
public class ProgressMessage {
    /**
     * STATE_FAILED 操作失败
     */
    public static final int STATE_FAILED = -1;
    /**
     * STATE_NOT_START 操作还未开始
     */
    public static final int STATE_NOT_START = 0;
    /**
     * STATE_STARTED 操作已经开始
     */
    public static final int STATE_STARTED = 1;
    /**
     * STATE_END 操作结束
     */
    public static final int STATE_END = 2;


    private int state = STATE_NOT_START;//状态(-1, 0, 1, 2)
    private int currentCount = 0;//当前执行到的数量
    private int totalCount = 0;//总数量
    private int progress = 0;//进度

    public void setTotalCount(int totalCount) {
        this.totalCount = totalCount;
    }

    public void setCurrentCount(int currentCount) {
        this.currentCount = currentCount;
        if (currentCount == 0) {
            progress = 0;
            this.state = STATE_STARTED;
            return;
        }
        if (totalCount == 0) {
            progress = 0;
            this.state = STATE_NOT_START;
            return;
        }
        if (currentCount >= totalCount) {
            progress = 100;
            this.state = STATE_END;
            return;
        }
        progress = currentCount * 100 / totalCount;
        this.state = STATE_STARTED;
    }

    /**
     * 重置进度
     */
    public void reset() {
        this.state = STATE_NOT_START;
        this.progress = 0;
        this.currentCount = 0;
        this.totalCount = 0;
    }


    public int getState() {
        return state;
    }

    public int getProgress() {
        return progress;
    }
}