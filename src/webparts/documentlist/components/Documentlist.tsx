import * as React from "react"
import { IDocumentlistProps } from "./IDocumentlistProps"
import { Sp } from "./Environments/Env"
import {
  FluentProvider,
  webLightTheme,
  Textarea,
  Title3,
  Label,
  useId,
  TextareaProps,
  Button,
  Body1,
} from "@fluentui/react-components"
import styles from "./Documentlist.module.scss"
import { Rating, RatingSize } from "office-ui-fabric-react/lib/Rating"

export default class Documentlist extends React.Component<any, any> {
  // Constructor declaration
  constructor(props) {
    super(props)
    this.state = {
      ratings: 0,
      currentuserData: undefined,
      comment: "",
    }
  }

  public async componentDidMount() {
    // Fetch the current user
    await Sp.currentUser()
      .then((res) => {
        console.log(res)
        this.setState({
          currentuserData: res,
        })
      })
      .catch((err) => {
        console.log(err)
      })
    console.log(this.state.currentuserData)
  }

  public render(): React.ReactElement<IDocumentlistProps> {
    const onRatingChange = (
      ev: React.FocusEvent<HTMLElement>,
      rating: number
    ): void => {
      this.setState({ ratings: rating })
    }

    const onTextChange: TextareaProps["onChange"] = (ev, data) => {
      this.setState({ comment: data.value })
    }

    const handleSubmit = async () => {
      if (this.state.comment !== "") {
        await Sp.lists
          .getByTitle("Feedback Ratings")
          .items.add({
            Title: this.state.currentuserData.Title,
            Comment: this.state.comment,
            Rating: this.state.ratings,
          })
          .then((res) => {
            alert(
              "Thanks for your valuable feedback! We appreciate your time and value your input as we work to improve our site experience."
            )
            this.setState({
              ratings: 0,
              comment: "",
            })
          })
          .catch((err) => {
            console.log(err)
          })
      } else {
        alert("Comment can't be empty")
      }
    }

    return (
      <FluentProvider theme={webLightTheme}>
        <section style={{ backgroundColor: "#fcfcfc" }}>
          <div style={{ textAlign: "center" }}>
            <Title3>Leave us a review</Title3>
          </div>
          <div className={styles.section1}>
            <Body1>
              How would you rate your level of satisfaction with this site?
            </Body1>
            <Rating
              min={0}
              max={5}
              size={RatingSize.Large}
              onChange={onRatingChange}
              rating={this.state.ratings}
            />
          </div>
          <div className={styles.section2}>
            <Label style={{ textAlign: "center" }}>
              Help us to make your experience even better. Share your comments
              with us:
            </Label>
            <div>
              <Textarea
                style={{
                  width: "354px",
                  height: "140px",
                }}
                onChange={onTextChange}
                resize='none'
                value={this.state.comment}
                placeholder='Type here..'
              />
            </div>
            <Button
              style={{ marginTop: "0.5rem" }}
              shape='circular'
              appearance='primary'
              iconPosition='after'
              onClick={handleSubmit}
            >
              Submit
            </Button>
          </div>
        </section>
      </FluentProvider>
    )
  }
}
