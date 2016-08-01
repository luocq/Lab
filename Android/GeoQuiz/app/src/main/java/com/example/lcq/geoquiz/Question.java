package com.example.lcq.geoquiz;

/**
 * Created by LCQ on 2016/7/29.
 */
public class Question {
    private int mTextResId;
    private  boolean mAnswerId;

    public int getTextResId() {
        return mTextResId;
    }

    public void setTextResId(int textResId) {
        mTextResId = textResId;
    }

    public boolean isAnswerId() {
        return mAnswerId;
    }

    public void setAnswerId(boolean answerId) {
        mAnswerId = answerId;
    }

    public Question(int textResId, Boolean answerId){
        this.mTextResId=textResId;

        this.mAnswerId=answerId;
    }
}
