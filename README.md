# CSV::Excel

CSV::Excel extends Ruby's `csv` stdlib to generate CSV file for Microsoft Excel.
This converts...

- Time/Date/DateTime object to Excel's date/time format (`yyyy/mm/dd`, `yyyy/mm/dd hh:mm:ss`).
- String to `="foo"` form for avoiding unexpected type casting (eg. Zero leading phone number, hyphen delimited street address).

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'csv-excel'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install csv-excel

## Usage

```ruby
require "csv/excel"

buffer = ""

csv = CSV.new(buffer, for_excel: true)
csv << ["012", "12", "🍣", 123, Date.new(2019, 4, 1), Time.new(2019, 4, 1, 12, 34, 56)]

# Excel reads csv as Shift_JIS by default.
# Please add BOM to indicate UTF-8 encoding.
File.write("out.csv", CSV::Excel::UTF8BOM + buffer)

# Open Finder/Explorer, and double-click `out.csv`!
```

## Development

After checking out the repo, run `bin/setup` to install dependencies. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/labocho/csv-excel.

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).
