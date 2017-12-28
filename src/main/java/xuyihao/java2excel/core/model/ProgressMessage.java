package xuyihao.java2excel.core.model;

import org.openxmlformats.schemas.drawingml.x2006.main.CTTextListStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Xuyh at 2016/07/26 上午 11:15.
 *
 * 进度通知（包括生成文件进度、导入文件进度）
 */
public class ProgressMessage {

    /**
     * 状态量
     *
     * STATE_FAILED 操作失败
     * STATE_NOT_START 操作还未开始
     * STATE_STARTED 操作已经开始
     * STATE_END 操作结束
     */
    public static final int STATE_FAILED = -1;
    public static final int STATE_NOT_START = 0;
    public static final int STATE_STARTED = 1;
    public static final int STATE_END = 2;

    /**
     * state 状态(-1, 0, 1, 2)
     * currentCount 操作的进度值
     * totalCount 当前操作的计数量进行到的位置
     * progress 操作的总计数量
     * detailMessage 操作进度详细情况描述
     */
    private int state = STATE_NOT_START;
    private int currentCount = 0;
    private int totalCount = 0;
    private int progress = 0;
    private String detailMessage = "";
    private List<String> warningMessageList = new ArrayList<String>();

    /**
     * 重置进度
     */
    public void reset(){
        setState(STATE_NOT_START);
        this.progress = 0;
        this.currentCount =0;
        this.totalCount = 0;
        this.detailMessage = "";
    }

    /**
     * 获取操作状态
     *
     * @return
     */
    public int getState() {
        return state;
    }

    /**
     * 设置操作状态
     *
     * @param state
     */
    public void setState(int state) {
        this.state = state;
    }

    /**
     * 操作未开始
     */
    public void stateNotStart(){
        this.setState(STATE_NOT_START);
    }

    /**
     * 操作开始
     */
    public void stateStarted(){
        this.setState(STATE_STARTED);
    }

    /**
     * 操作结束
     */
    public void stateEnd(){
        this.setState(STATE_END);
        this.setProgress(100);
    }

    /**
     * 操作失败
     */
    public void stateFailed(){
        this.setState(STATE_FAILED);
		this.setProgress(100);
    }

    /**
     * 当前操作位置增一
     *
     */
    public void addCurrentCount(){
        currentCount++;
        setCurrentCount(currentCount);
    }

    /**
     * 获取当前操作到的位置
     * @return
     */
    public int getCurrentCount() {
        return currentCount;
    }

    /**
     * 设置当前操作进行到的位置
     *
     * @param currentCount
     */
    public void setCurrentCount(int currentCount) {
        this.currentCount = currentCount;
        if(currentCount == 0){
            progress = 0;
            setState(STATE_STARTED);
            return;
        }
        if(totalCount == 0){
            progress = 0;
            setState(STATE_NOT_START);
            return;
        }
        if(currentCount >= totalCount){
            progress = 100;
            setState(STATE_END);
            return;
        }
        progress = currentCount * 100 / totalCount;
        setState(STATE_STARTED);
    }

    /**
     * 获取操作总数
     *
     * @return
     */
    public int getTotalCount() {
        return totalCount;
    }

    /**
     * 设置操作总数
     *
     * @param totalCount
     */
    public void setTotalCount(int totalCount) {
        this.totalCount = totalCount;
    }

    /**
     * 获取操作进度
     *
     * @return
     */
    public int getProgress() {
        return progress;
    }

    /**
     * 设置进度
     *
     * @param progress（1-100）
     */
    public void setProgress(int progress) {
        this.progress = progress;
    }

    /**
     * 获取进度的详细信息
     *
     * @return
     */
    public String getDetailMessage() {
        return detailMessage;
    }

    /**
     * 设置进度的详细信息
     *
     * @param detailMessage
     */
    public void setDetailMessage(String detailMessage) {
        this.detailMessage = detailMessage;
        System.out.println(detailMessage);
    }

    public List<String> getWarningMessageList(){
        return this.warningMessageList;
    }

    public void setWarningMessageList(List<String> warningMessageList){
        this.warningMessageList = warningMessageList;
    }

    public void addWarningMessage(String warningMessage){
        this.setDetailMessage("Warning: " + warningMessage);
        this.warningMessageList.add(warningMessage);
    }
}