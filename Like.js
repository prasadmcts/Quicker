import cx from 'classnames';
import { Component } from 'react';

export default class LikeButton extends Component {
    constructor(props){
        super(props);
        this.state = {
            likes: 100,
            isClicked: false
        };
    }
    handleLikes= () => {
        if(!this.state.isClicked)
        {
            this.setState(prevState => ({
                likes: prevState.likes + 1,
                isClicked: true
            }))
        }
        else {
            this.setState(prevState => ({
                likes: prevState.likes - 1,
                isClicked: false
            }))
        }
    }

    render() {
        return (
            <>
                <div>
                    <button className={this.state.isClicked ? 'liked' : 'like-button'}  onClick= {this.handleLikes}> Like | {this.state.likes}</button>
                </div>
                <style>{`
                    .like-button {
                        font-size: 1rem;
                        padding: 5px 10px;
                        color:  #585858;
                    }
                   .liked {
                        font-weight: bold;
                        color: #1565c0;
                   }
                `}</style>
            </>
        );
    }
}


==========================
import React, {useState} from "react";
import "./style.css";

export default function App() {
  const [countObj, setCountObj] = useState({count: 100, isClicked: false});
  const handleLike = () => {
    if (countObj.isClicked) {
      setCountObj({count: countObj.count-1, isClicked: false});
    } else {
      setCountObj({count: countObj.count+1, isClicked: true});
    }
  }
  return (
    <div>
      Count: {countObj.count}
      <button onClick={handleLike}>LIKE</button>
    </div>
  );
}

=======================
 import { useState } from "react";
import "./styles.css";

export default function App() {
  const [likes, setLikes] = useState({
    count: 100,
    isLiked: false
  });

  const updateLikes = () => {
    setLikes((data) => {
      const newCount = data.isLiked ? data.count - 1 : data.count + 1;
      return {
        count: newCount,
        isLiked: !data.isLiked
      };
    });
  };

  return (
    <div className="App">
      <LikeButton {...likes} updateLikes={updateLikes} />
    </div>
  );
}

const LikeButton = ({ count, updateLikes }) => {
  return <button onClick={() => updateLikes()}>{count} likes</button>;
};
   
