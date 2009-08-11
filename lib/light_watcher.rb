require 'rubygems'

#Knows when the traffic light changes.
# parses ZoneMinder log file.
class LightWatcher
  attr_accessor :zone_minder_log_file
end


if $0 == __FILE__
  begin
    require 'redgreen';
  rescue LoadError;
  end
  require 'test/unit'
  require 'rr'

  class Test::Unit::TestCase
    include RR::Adapters::TestUnit
  end

  class TestLightWatcher < Test::Unit::TestCase
    def setup
      @it = LightWatcher.new
    end

    def test_should_respond_to_zone_minder_log
      assert @it.respond_to?(:zone_minder_log_file)
    end

  end
end

