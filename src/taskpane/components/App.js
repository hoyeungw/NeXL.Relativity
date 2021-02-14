import * as React                 from "react"
import { ButtonType }             from "office-ui-fabric-react"
import { PrimaryButton }          from "office-ui-fabric-react"
import Header                     from "./Header"
import HeroList, { HeroListItem } from "./HeroList"
import Progress                   from "./Progress"
import { fluoMatrix }             from "@palett/fluo-matrix"
import { POINTWISE }              from '@vect/enum-matrix-directions'
import { FRESH, PLANET }          from '@palett/presets'
import { mapper }                 from '@vect/matrix'
import { deco }                   from '@spare/deco'
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context)
    this.state = {
      listItems: []
    }
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    })
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        const range = context.workbook.getSelectedRange()
        range.load(["address", "values", "rowCount", "columnCount"]) // Read the range properties
        await context.sync()
        console.log(`The range address was ${range.address}.`)
        let values = range.values
        let colors = fluoMatrix(values, { direct: POINTWISE, presets: [PLANET, FRESH] })
        console.log(JSON.stringify(colors, null, 2))
        let height = range.rowCount, width = range.columnCount
        console.log(`H ${height}, W ${width}`)
        // Update the fill color
        // range.format.fill.color = "yellow"
        for (let i = 0; i < height; i++) {
          for (let j = 0; j < width; j++) {
            // range.getCell(i, j).format.fill.color = colors[i][j]
          }
        }
        // range.values = mapper(values, x => x + 1)
        range.format.fill.color = "#909090"
      })
    } catch (error) {
      console.error(error)
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png"
                  message="Please sideload your addin to see app body."/>
      )
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome"/>
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <PrimaryButton
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </PrimaryButton>
        </HeroList>
      </div>
    )
  }
}
