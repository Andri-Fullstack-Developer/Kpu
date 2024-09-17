<?php

namespace App\Events;

use Illuminate\Broadcasting\Channel;
use Illuminate\Broadcasting\InteractsWithSockets;
use Illuminate\Foundation\Events\Dispatchable;
use Illuminate\Queue\SerializesModels;

class VoteUpdated
{
    use Dispatchable, InteractsWithSockets, SerializesModels;

    public $voteData;

    /**
     * Create a new event instance.
     */
    public function __construct($voteData)
    {
        $this->voteData = $voteData;
    }

    /**
     * Get the channels the event should broadcast on.
     *
     * @return array<int, \Illuminate\Broadcasting\Channel>
     */
    public function broadcastOn()
    {
        return new Channel('dashboard');
    }
}