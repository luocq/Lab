package com.example.lcq.geoquiz;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.TextView;
import android.widget.Toast;

public class QuizActivity extends AppCompatActivity {
    private Button mTrueButton;
    private Button mFalseButton;
    private Button mNextButton;
    private TextView mQuestionTextView;
    private static final String TAG="QuizActivity";

    private Question[] mQuestionBank=new Question[]{
        new Question(R.string.qustion_africa,true),
            new Question(R.string.qustion_americas,true),
            new Question(R.string.qustion_asia,true),
            new Question(R.string.qustion_mideast,true),
            new Question(R.string.qustion_oceans,true),
    };

    private int mCurrentIndex=0;

    private  void UpdateQuestion(){
        int question = mQuestionBank[mCurrentIndex].getTextResId();
        mQuestionTextView.setText(question);
    }
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_quiz);
        Log.d(TAG,"启动");

        mQuestionTextView = (TextView)findViewById(R.id.question_text_view);
        mTrueButton = (Button) findViewById(R.id.true_button);
        mFalseButton= (Button)findViewById(R.id.false_button);
        mNextButton = (Button)findViewById(R.id.next_button);

        UpdateQuestion();

        mTrueButton.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                //TODO
                //Toast.makeText(QuizActivity.this,R.string.correct_toast,Toast.LENGTH_SHORT).show();
                CheckAnswer(true);
            }
        });

        mFalseButton.setOnClickListener(new View.OnClickListener(){
            @Override
            public  void onClick(View view){
                CheckAnswer(false);
            }
        });
        mNextButton.setOnClickListener(new View.OnClickListener(){
            @Override
            public void onClick(View view) {
                mCurrentIndex=(mCurrentIndex+1)%mQuestionBank.length;
                UpdateQuestion();
                int question = mQuestionBank[mCurrentIndex].getTextResId();
                mQuestionTextView.setText(question);
            }
        });
    }


    @Override
    protected void onStart()
    {
        super.onStart();
        Log.d(TAG,"onStart()");
    }
    @Override
    protected void onPause()
    {
        super.onPause();
        Log.d(TAG,"onPause()");
    }
    @Override
    protected void onStop()
    {
        super.onStop();
        Log.d(TAG,"onStop() called");
    }
    @Override
    protected void onDestroy()
    {
        super.onDestroy();
        Log.d(TAG,"onDestroy() called");
    }
    private void CheckAnswer(boolean userPressedTrue)
    {
        boolean answerIsTrue = mQuestionBank[mCurrentIndex].isAnswerTrue();

        int messageResId=0;
        if (userPressedTrue==answerIsTrue){
            messageResId=R.string.correct_toast;
        }
        else
        {
            messageResId=R.string.incorrect_toast;
        }

        Toast.makeText(this,messageResId,Toast.LENGTH_SHORT).show();
    }
}
