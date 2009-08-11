require 'proxr'

class VideoController
  attr_accessor :pro_xr

  def initialize
    @pro_xr = ProXR.new
  end

  def default_bank
    1
  end

  def relay_number(relay_indicator)
    case relay_indicator
      when :recording
        0
      else
        raise "That relay_indicator (#{relay_indicator}) is not supported."
    end
  end

  def relay_off(relay_number)
    @pro_xr.relay_off(relay_number, default_bank)
  end

  def relay_on(relay_number)
    @pro_xr.relay_on(relay_number, default_bank)
  end

  def relay_on?(relay_number)
    @pro_xr.relay_on?(relay_number, default_bank)
  end

  def record_violation(lane_number)
    tell_overhead_camera_to_record_violation
    start_recording(lane_number)
    sleep violation_recording_duration
    stop_recording(lane_number)
  end

  def start_recording(lane_number)
    relay_on(lane_number-1) #lane is one-based, relay is zero-based
  end

  def stop_recording(lane_number)
    relay_off(lane_number-1) #lane is one-based, relay is zero-based
  end

  def tell_overhead_camera_to_record_violation
    #TODO: external controller
  end

  def violation_recording_duration
    12.seconds
  end

end


if $0 == __FILE__
  require 'test/unit'
  require 'rr'

  class Test::Unit::TestCase
    include RR::Adapters::TestUnit
  end

  class TestVideoController < Test::Unit::TestCase
    def setup
      @it = VideoController.new
    end

    def test_start_recording_should_turn_relay_on
      lane_number = 2
      @it.start_recording(lane_number)
      assert @it.relay_on?(lane_number-1)
    end

    def test_stop_recording_should_turn_relay_off
      lane_number = 3
      @it.start_recording(lane_number)
      assert @it.relay_on?(lane_number-1)
      @it.stop_recording(lane_number)
      assert !@it.relay_on?(lane_number-1)
    end
  end

  class TestVideoControllerRecordViolation < Test::Unit::TestCase
    def setup
      @it = VideoController.new
      stub(@it).tell_overhead_camera_to_record_violation
      stub(@it).start_recording(numeric)
      stub(@it).violation_recording_duration {0}
      stub(@it).stop_recording(numeric)
    end

    def test_record_violation_should_tell_the_overhead_camera_to_record_violation
      mock(@it).tell_overhead_camera_to_record_violation
      @it.record_violation(1)
    end

    def test_record_violation_should_tell_the_lane_to_start_recording
      lane_number = 2
      mock(@it).start_recording(lane_number)
      @it.record_violation(lane_number)
    end

    def test_record_violation_should_tell_the_lane_to_stop_recording_after_waiting
      mock(@it).sleep(@it.violation_recording_duration)
      lane_number = 1
      mock(@it).stop_recording(lane_number)
      @it.record_violation(lane_number)
    end


    def test_should_record_violation_if_vehicle_enters_violation_zone_and_violation_is_possible #assumes channel only receives  violations are possible.
      
    end
  end
end
