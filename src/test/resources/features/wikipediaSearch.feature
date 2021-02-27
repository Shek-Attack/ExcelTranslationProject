Feature: Wikipedia search feature
  As a user, I should be able to search terms and see related results

    #windows: alt + enter
    #mac: option + enter

  Scenario: Wikipedia search title verification
    Given User is on Wikipedia home page
    When User types "Steve Jobs" in the wiki search box
    And User clicks wiki search button
    Then User sees "Steve Jobs" is in the wiki title

  @wiki
  Scenario: Wikipedia search header verification
    Given User is on Wikipedia home page
    When User types "Steve Jobs" in the wiki search box
    And User clicks wiki search button
    Then User sees "Steve Jobs" in the main header


  Scenario Outline: Wikipedia search header verification
    Given User is on Wikipedia home page
    When User types "<searchValue>" in the wiki search box
    And User clicks wiki search button
    Then User sees "<expectedMainHeader>" in the main header
    Then User sees "<expectedTitle>" is in the wiki title

    Examples: search terms we are going to search in wikipedia
      | searchValue        | expectedMainHeader | expectedTitle      |
      | Steve Jobs         | Steve Jobs         | Steve Jobs         |
      | Maryam Mirzakhani  | Maryam Mirzakhani  | Maryam Mirzakhani  |
      | Mahmud al-Kashgari | Mahmud al-Kashgari | Mahmud al-Kashgari |
      | Elon Musk          | Elon Musk          | Elon Musk          |
      | Bill Gates         | Bill Gates         | Bill Gates         |
      | Muhammad Ali       | Muhammad Ali       | Muhammad Ali       |
      | Marie Curie        | Marie Curie        | Marie Curie        |

