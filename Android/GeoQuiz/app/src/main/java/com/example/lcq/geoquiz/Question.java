package com.example.lcq.geoquiz;

/**
 * Created by LCQ on 2016/7/29.
 */
public class Question {
    private int mTextResId;
    private  boolean mAnswerTrue;

    public int getTextResId() {
        return mTextResId;
    }

    public void setTextResId(int textResId) {
        mTextResId = textResId;
    }

    public boolean isAnswerTrue() {
        return mAnswerTrue;
    }

    public void setAnswerTrue(boolean answerTrue) {
        mAnswerTrue = answerTrue;
    }

    public Question(int textResId, Boolean answerTrue){
        this.mTextResId=textResId;

        this.mAnswerTrue=answerTrue;
    }
}
