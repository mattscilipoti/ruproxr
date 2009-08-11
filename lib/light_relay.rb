require 'proxr'

class LightRelay
  attr_accessor :pro_xr

  def initialize
    @pro_xr = ProXR.new
  end

  def default_bank
    1
  end

  def relay_number(relay_indicator)
    case relay_indicator
      when :green
        0
      when :yellow
        1
      when :red
        2
      else
        raise "That relay_indicator (#{relay_indicator}) is not supported."
    end
  end

  def relay_off(relay_indicator)
    @pro_xr.relay_off(relay_number(relay_indicator), default_bank)
  end

  def relay_on(relay_indicator)
    @pro_xr.relay_on(relay_number(relay_indicator), default_bank)
  end

  def relay_on?(relay_indicator)
    @pro_xr.relay_on?(relay_number(relay_indicator), default_bank)
  end

  def shift_relays(relay_to_turn_off, relay_to_turn_on)
    relay_off(relay_to_turn_off)
    relay_on(relay_to_turn_on)
  end
end


if $0 == __FILE__
  require 'test/unit'
  class TestLightRelay < Test::Unit::TestCase
    def setup
      @it = LightRelay.new
    end

    def test_green
      assert_equal 0, @it.relay_number(:green)
    end

    def test_yellow
      assert_equal 1, @it.relay_number(:yellow)
    end

    def test_red
      assert_equal 2, @it.relay_number(:red)
    end

    def test_red_to_green
      @it.shift_relays(:red, :green)
      assert @it.relay_on?(:green)
      assert !@it.relay_on?(:red)
    end
  end
end
