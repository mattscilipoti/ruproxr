require 'rubygems'
require 'serialport'
require 'timeout'

class ProXR < SerialPort
  COMMAND_MODE = 254
  READ_VOLTAGE_8_BITS = 150
  RELAY_OFF    = 100
  RELAY_ON     = 108
  RELAY_STATUS = 116
  REPORTING    = 27
  SUCCESS      = 85

  #SerialPort uses ::new instead of initialize.
  def ProXR::new(port = "/dev/ttyS0", *params)
    default_params = [115200]
    default_params.each_with_index {|param, index| params[index] = param unless params[index] }
    super port, *params
#    read_timeout = 2 #does this do anything?
  end

  def self.read_voltage_command(channel, data_bits = 8)
    base_cmd =
      case data_bits
        when 8
          READ_VOLTAGE_8_BITS
        else
          raise "That data_bit value (#{data_bits}) is not supported."
      end
    base_cmd + channel
  end

  #Does this channel have any voltage?
  def has_voltage?(channel)
    read_voltage(channel) > 0 ? true : false
  end

  #used by read_voltage
  def max_channel_reading(data_bits)
    case data_bits
      when 8
        max_reading_for_8_bit = 255
      else
        raise "That data_bit value (#{data_bits}) is not supported."
    end
  end

  #used by read_voltage
  def max_channel_voltage
    5
  end

  #puts the device in Command Mode, then
  #runs each command,
  #returning the final value from ProXR
  #Has Timeout.
  def process_commands(*cmds)
    Timeout::timeout(1) do
      send_command(COMMAND_MODE)
      cmds.each {|cmd| send_command(cmd) }
      getc
    end
  end

  #Command is expected to return SUCCESS
  # raises error if not SUCCESS
  def process_commands!(*cmds)
    result = process_commands(*cmds)
    raise CommandError, "Expected #{SUCCESS}, received #{result}" unless result == SUCCESS
    result
  end

  #Reads the voltage from the specified channel.
  #Defaults to 8 bit.
  # Supports:
  # * 8 bit
  def read_voltage(channel, data_bits = 8)
    ad_voltage = process_commands(ProXR.read_voltage_command(channel))

    voltage = ad_voltage / voltage_conversion_factor(data_bits)

    #special case
    voltage = 0 if ad_voltage == max_channel_reading(data_bits)
    voltage
  end

  #Activiate Relay
  def relay_on(relay_number, bank_number)
    cmd = RELAY_ON
    relay_on_cmd = (cmd + relay_number)
    process_commands! relay_on_cmd, bank_number
  end

  #Is this thing on?
  def relay_on?(relay_number, bank_number)
    relay_status(relay_number, bank_number) == 1
  end

  #Deactivate Relay
  def relay_off(relay_number, bank_number)
    relay_off_cmd = (RELAY_OFF + relay_number)
    process_commands! relay_off_cmd, bank_number
  end

  #What is the status (0/1) of this relay?
  def relay_status(relay_number, bank_number)
    relay_status_cmd = (RELAY_STATUS + relay_number)
    process_commands relay_status_cmd, bank_number
  end

  #Puts device in Reporting Mode.
  def reporting_mode
    process_commands!(REPORTING)
  end

  #Actually sends the command to the device.
  #Converts the command to chr prior to sending.
  def send_command(cmd)
    write cmd.chr
  end

  #Scan all
  def show_all_voltages
#    puts "Ch:Volt"
    (0..7).collect do |channel|
      "#{channel}:#{read_voltage(channel)}"
    end
  end

  #used by read voltage
  def voltage_conversion_factor(data_bits)
    max_channel_reading(data_bits)/max_channel_voltage
  end


  class CommandError < StandardError

  end
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


  #TODO: separate into "live" vs behavior testa (do we interact with the device correctly?)
  class TestProXR_basic_command_processing < Test::Unit::TestCase
    def setup
      @it = ProXR.new
      stub(@it).getc { ProXR::SUCCESS }
    end

    def test_send_command_write_converted_chr
      mock(@it).write(1.chr)
      @it.send_command(1)
    end

    def test_process_commands_should_start_command_mode_prior_to_sending_command
      mock(@it).send_command(ProXR::COMMAND_MODE)
      mock(@it).send_command(1)

      @it.process_commands 1
    end

    def test_process_commands_bang_should_error_if_not_success
      stub(@it).process_commands { 0 }
      assert_raises ProXR::CommandError do
        @it.process_commands!
      end
    end
  end

  class TestProXR_commands < Test::Unit::TestCase
    def setup
      @it = ProXR.new
      stub(@it).send_command(anything) { ProXR::SUCCESS }
#      stub(@serial_port).relay_off(anything, anything) { ProXR::SUCCESS }
#      stub(@serial_port).relay_on(anything, anything) { ProXR::SUCCESS }
    end

    def test_has_voltage_without_voltage
      mock(@it).read_voltage(2) { 0 }
      assert !@it.has_voltage?(2)
    end

    def test_has_voltage_with_voltage
      mock(@it).read_voltage(2) { 1.4 }
      assert @it.has_voltage?(2)
    end

    def test_reporting_mode_should_send_the_reporting_command
      mock(@it).process_commands!(ProXR::REPORTING) { ProXR::SUCCESS }
      assert_equal ProXR::SUCCESS, @it.reporting_mode
    end

    def test_reporting_command_should_use_send_command_bang
      mock(@it).process_commands!(ProXR::REPORTING) { ProXR::SUCCESS }
      @it.reporting_mode
    end

    def test_read_voltage_process_the_read_voltage_command
      mock(@it).process_commands(ProXR::READ_VOLTAGE_8_BITS + 2) { 2.4 * @it.voltage_conversion_factor(8) }
      assert_equal 2.4, @it.read_voltage(2)
    end

    def test_relay_on_should_call_process_commands_bang
      mock(@it).process_commands!(anything, anything)
      @it.relay_on(1, 2)
    end

    def test_relay_on_should_process_the_relay_on_command
      mock(@it).process_commands!(ProXR::RELAY_ON + 1, 2) { ProXR::SUCCESS }
      assert_equal ProXR::SUCCESS, @it.relay_on(1, 2)
    end

    def test_relay_off_should_call_process_commands_bang
      mock(@it).process_commands!(anything, anything)
      @it.relay_off(1, 2)
    end

    def test_relay_off_should_process_the_relay_off_command
      mock(@it).process_commands!(ProXR::RELAY_OFF + 2, 3) { ProXR::SUCCESS }
      assert_equal ProXR::SUCCESS, @it.relay_off(2, 3)
    end

    def test_relay_status_should_the_relay_status_command
      mock(@it).process_commands(ProXR::RELAY_STATUS + 3, 1) { 1 }
      assert_equal 1, @it.relay_status(3, 1)
    end

    def test_relay_on_question_when_on_relay_on
      stub(@it).relay_status(anything, anything) { 1 }
      assert @it.relay_on?(0, 1)
    end

    def test_relay_off_question_when_on_relay_on
      stub(@it).relay_status(anything, anything) { 0 }
      assert (not @it.relay_on?(0, 1))
    end

  end

  unless ENV['DEVICE'] =~ /true/i
    puts "Skipping tests against actual device.  Use `DEVICE=true ruby proxr.rb` to enable these tests."
  else
    class TestProXR_Live < Test::Unit::TestCase
      def setup
        @serial_port = ProXR.new
      end

      def test_should_indicate_it_is_in_reporting_mode
        assert_equal ProXR::SUCCESS, @serial_port.reporting_mode
      end

      def test_voltage_at_0_1_should_be_zero
        assert_equal 0, @serial_port.read_voltage(0)
      end

      def test_relay_on_0_1
        assert_equal ProXR::SUCCESS, @serial_port.relay_on(0, 1)
      end

      def test_relay_off_0_1
        assert_equal ProXR::SUCCESS, @serial_port.relay_off(0, 1)
      end

      def test_relay_status_for_on
        @serial_port.relay_on(0, 1)
        assert_equal 1, @serial_port.relay_status(0, 1)
      end

      def test_relay_status_for_off
        @serial_port.relay_off(0, 1)
        assert_equal 0, @serial_port.relay_status(0, 1)
      end

      def test_when_on_relay_on?
        @serial_port.relay_on(0, 1)
        assert @serial_port.relay_on?(0, 1)
      end

      def test_when_off_relay_on?
        @serial_port.relay_off(0, 1)
        assert !@serial_port.relay_on?(0, 1)
      end

      def test_relay_on_should_raise_error_if_result_is_not_SUCCESS
        mock(@serial_port).getc { 0 }

        assert_raise ProXR::CommandError do
          @serial_port.relay_on(0, 1)
        end
      end
    end
  end

end
