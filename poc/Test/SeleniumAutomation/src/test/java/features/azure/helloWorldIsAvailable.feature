@SMOKE
Feature: As an infrastructure team, after a deployment I want to assure that the Hello world site is running.

   @WEB @HelloWorld
   Scenario: Check that Hello World is up and running
     Given I navigate to Hello World
     When I check the body text
     Then I expect the hello world landing page to say "Hi"





